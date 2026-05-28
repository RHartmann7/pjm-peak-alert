"""
PJM 5CP Peak Load Alert
=======================

Pulls today's PJM RTO load forecast from PJM Data Miner 2, tiers it against
5CP risk thresholds, and emails a daily summary plus an elevated warning when
forecast peak load suggests a likely Peak Load Contribution (PLC) setting hour.

Designed for ABCO / Maineville Storage and any other PJM C&I customer that
wants to manage Peak Load Contribution by curtailing load during the five
highest PJM RTO peak hours each summer.

USAGE
-----
1.  Get a free PJM Data Miner 2 subscription key:
        https://dataminer2.pjm.com/  ->  My Profile  ->  Subscriptions
2.  Get a Microsoft 365 mailbox + an app password (or use a service account
    with SMTP AUTH enabled).
3.  Set the environment variables below (copy from .env.example or set in
    your shell / Task Scheduler / cron).
4.  Run:  python pjm_peak_alert.py
    Optional flags:
        --dry-run       Print the email body to stdout instead of sending.
        --csv PATH      Use a local CSV instead of the live API (for testing).
        --date YYYY-MM-DD   Override the target date (default: today, ET).

SCHEDULING
----------
Run once per day in the morning so you have time to plan curtailment.
Windows Task Scheduler (recommended trigger: 07:00 ET, Mon-Sun, Jun 1 - Sep 30):
    Action:  python.exe  C:\\path\\to\\pjm_peak_alert.py
Linux cron:
    0 7 * 6-9 *  /usr/bin/python3 /opt/pjm/pjm_peak_alert.py

ENVIRONMENT VARIABLES
---------------------
    PJM_API_KEY          PJM Data Miner 2 subscription key (required for live mode)
    SMTP_HOST            Default: smtp.office365.com
    SMTP_PORT            Default: 587
    SMTP_USER            Your M365 email address (sender)
    SMTP_PASS            App password or account password
    MAIL_FROM            From address (usually same as SMTP_USER)
    MAIL_TO              Comma-separated recipient list
    MAIL_REPLY_TO        Optional reply-to address
    ALERT_THRESHOLDS     Optional override "watch,warning,curtail" in MW
                         (default: 140000,150000,158000)
"""

from __future__ import annotations

import argparse
import csv
import os
import smtplib
import ssl
import sys
import textwrap
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from typing import Iterable
from zoneinfo import ZoneInfo

try:
    import requests
except ImportError:
    requests = None  # only required for live mode


# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------

PJM_API_BASE = "https://api.pjm.com/api/v1"
PJM_FORECAST_FEED = "load_frcstd_7_day"   # 7-day hourly load forecast
DEFAULT_THRESHOLDS_MW = (140_000, 150_000, 158_000)   # watch, warning, curtail
EASTERN = ZoneInfo("America/New_York")


@dataclass
class Tier:
    name: str
    color: str
    emoji: str
    action: str


TIERS = {
    "GREEN":   Tier("GREEN",   "#2E7D32", "🟢", "No action. Operate normally."),
    "YELLOW":  Tier("YELLOW",  "#F9A825", "🟡", "Monitor. Be ready to curtail if PJM forecast climbs through the day."),
    "ORANGE":  Tier("ORANGE",  "#EF6C00", "🟠", "Likely 5CP candidate. Pre-cool building by 11 a.m. ET; review the curtailment SOP; flag operators."),
    "RED":     Tier("RED",     "#C62828", "🔴", "HIGH probability 5CP day. Execute curtailment plan 2 p.m. – 7 p.m. ET. Shed non-essential HVAC, defer compressed-air/charging loads, stagger machine starts."),
}


def tier_for_mw(peak_mw: float, thresholds=DEFAULT_THRESHOLDS_MW) -> Tier:
    watch, warning, curtail = thresholds
    if peak_mw >= curtail:
        return TIERS["RED"]
    if peak_mw >= warning:
        return TIERS["ORANGE"]
    if peak_mw >= watch:
        return TIERS["YELLOW"]
    return TIERS["GREEN"]


# ----------------------------------------------------------------------------
# Data acquisition
# ----------------------------------------------------------------------------

def parse_pjm_dt(value: str) -> datetime:
    """PJM Data Miner emits ISO-ish strings; the CSV emits M/D/YYYY H:MM."""
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%m/%d/%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    # Fall through to fromisoformat for safety
    return datetime.fromisoformat(value)


def fetch_forecast_live(api_key: str, target_date) -> list[dict]:
    """Hit Data Miner 2 for the latest 7-day forecast and filter to target_date."""
    if requests is None:
        raise RuntimeError("The `requests` package is required for live mode. `pip install requests`.")

    url = f"{PJM_API_BASE}/{PJM_FORECAST_FEED}"
    # Data Miner uses EPT for forecast_hour_beginning_ept. Pull two days to be
    # safe across midnight boundaries, then filter locally.
    start_ept = datetime.combine(target_date, datetime.min.time())
    end_ept   = datetime.combine(target_date, datetime.max.time())
    params = {
        "rowCount": 5000,
        "forecast_area": "RTO",
        "forecast_hour_beginning_ept": f"{start_ept:%Y-%m-%dT%H:%M:%S} to {end_ept:%Y-%m-%dT%H:%M:%S}",
    }
    headers = {
        "Ocp-Apim-Subscription-Key": api_key,
        "Accept": "application/json",
    }
    resp = requests.get(url, params=params, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("items", data) if isinstance(data, dict) else data


def fetch_forecast_csv(path: str, target_date) -> list[dict]:
    """Local CSV reader for offline / dry-run testing."""
    rows: list[dict] = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row.get("forecast_area", "").upper() != "RTO":
                continue
            try:
                hour_ept = parse_pjm_dt(row["forecast_hour_beginning_ept"])
            except Exception:
                continue
            if hour_ept.date() != target_date:
                continue
            row["_hour_ept"] = hour_ept
            row["_evaluated_ept"] = parse_pjm_dt(row["evaluated_at_ept"])
            row["_load_mw"] = float(row["forecast_load_mw"])
            rows.append(row)
    return rows


# ----------------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------------

@dataclass
class DayPeak:
    target_date: object
    peak_mw: float
    peak_hour_ept: datetime
    evaluated_at_ept: datetime
    all_hours: list[tuple[datetime, float]]
    tier: Tier


def latest_forecast_per_hour(rows: Iterable[dict]) -> dict[datetime, dict]:
    """Each hour can have multiple forecast revisions. Keep the most recent
    `evaluated_at_ept` per `forecast_hour_beginning_ept`."""
    latest: dict[datetime, dict] = {}
    for r in rows:
        hour = r.get("_hour_ept") or parse_pjm_dt(r["forecast_hour_beginning_ept"])
        evald = r.get("_evaluated_ept") or parse_pjm_dt(r["evaluated_at_ept"])
        mw = r.get("_load_mw")
        if mw is None:
            mw = float(r["forecast_load_mw"])
        cur = latest.get(hour)
        if cur is None or evald > cur["_evaluated_ept"]:
            latest[hour] = {"_hour_ept": hour, "_evaluated_ept": evald, "_load_mw": mw}
    return latest


def compute_day_peak(rows, target_date, thresholds=DEFAULT_THRESHOLDS_MW) -> DayPeak | None:
    latest = latest_forecast_per_hour(rows)
    if not latest:
        return None
    by_hour = sorted(latest.values(), key=lambda r: r["_hour_ept"])
    peak = max(by_hour, key=lambda r: r["_load_mw"])
    return DayPeak(
        target_date=target_date,
        peak_mw=peak["_load_mw"],
        peak_hour_ept=peak["_hour_ept"],
        evaluated_at_ept=peak["_evaluated_ept"],
        all_hours=[(r["_hour_ept"], r["_load_mw"]) for r in by_hour],
        tier=tier_for_mw(peak["_load_mw"], thresholds),
    )


# ----------------------------------------------------------------------------
# Email
# ----------------------------------------------------------------------------

def render_email(dp: DayPeak, thresholds=DEFAULT_THRESHOLDS_MW) -> tuple[str, str, str]:
    """Return (subject, plain_text_body, html_body)."""
    date_str = dp.target_date.strftime("%A, %B %d, %Y")
    peak_gw = dp.peak_mw / 1000.0
    peak_hour_str = dp.peak_hour_ept.strftime("%-I:%M %p ET") if os.name != "nt" \
                    else dp.peak_hour_ept.strftime("%#I:%M %p ET")
    evaluated_str = dp.evaluated_at_ept.strftime("%Y-%m-%d %H:%M ET")

    subject = f"[{dp.tier.name}] PJM Peak Forecast {date_str}: {peak_gw:,.1f} GW @ {peak_hour_str}"

    watch, warning, curtail = (t / 1000 for t in thresholds)

    hourly_lines = [
        f"  {h.strftime('%-H' if os.name != 'nt' else '%#H'):>2}:00  {mw/1000:>6.1f} GW"
        for h, mw in dp.all_hours
    ]

    plain = textwrap.dedent(f"""\
        PJM RTO Load Forecast — {date_str}

        Tier:           {dp.tier.emoji} {dp.tier.name}
        Forecast peak:  {peak_gw:,.1f} GW (= {dp.peak_mw:,.0f} MW)
        Peak hour:      {peak_hour_str}
        Forecast issued: {evaluated_str}

        ACTION
        ------
        {dp.tier.action}

        TIER THRESHOLDS (RTO forecast peak, GW)
        ---------------------------------------
        Green   <  {watch:.0f}
        Yellow  ≥  {watch:.0f}
        Orange  ≥  {warning:.0f}
        Red     ≥  {curtail:.0f}

        HOURLY FORECAST (latest revision, hour beginning, ET)
        -----------------------------------------------------
        {chr(10).join(hourly_lines)}

        Source: PJM Data Miner 2, feed `{PJM_FORECAST_FEED}`.
        About 5CP / PLC: capacity charges on your Choice (Dynegy) bill scale
        with your Peak Load Contribution, set by your average demand during
        the five highest PJM RTO peak hours each summer (Jun 1 – Sep 30).
        Curtailing 50–75 kW during those hours can save $6K–$16K/year.
    """)

    rows_html = "".join(
        f"<tr><td style='padding:2px 10px'>{h.strftime('%H:%M')}</td>"
        f"<td style='padding:2px 10px;text-align:right'>{mw/1000:,.1f}</td></tr>"
        for h, mw in dp.all_hours
    )

    html = f"""\
<!DOCTYPE html><html><body style="font-family:Calibri,Arial,sans-serif;color:#222">
<h2 style="color:{dp.tier.color};margin-bottom:4px">{dp.tier.emoji} {dp.tier.name} — PJM Peak Forecast</h2>
<div style="color:#666;margin-bottom:16px">{date_str}</div>

<table style="border-collapse:collapse;margin-bottom:16px">
  <tr><td style="padding:4px 12px;color:#555">Forecast peak</td>
      <td style="padding:4px 12px;font-weight:bold;font-size:18px">{peak_gw:,.1f} GW <span style="color:#888;font-weight:normal">({dp.peak_mw:,.0f} MW)</span></td></tr>
  <tr><td style="padding:4px 12px;color:#555">Peak hour</td>
      <td style="padding:4px 12px"><b>{peak_hour_str}</b></td></tr>
  <tr><td style="padding:4px 12px;color:#555">Forecast issued</td>
      <td style="padding:4px 12px">{evaluated_str}</td></tr>
</table>

<div style="padding:12px 16px;border-left:4px solid {dp.tier.color};background:#FAFAFA;margin-bottom:16px">
  <b>Action:</b> {dp.tier.action}
</div>

<h3 style="margin-bottom:4px">Tier thresholds (RTO peak, GW)</h3>
<table style="border-collapse:collapse;margin-bottom:16px">
  <tr><td style="padding:2px 12px;color:{TIERS['GREEN'].color}">🟢 Green</td><td style="padding:2px 12px">&lt; {watch:.0f}</td></tr>
  <tr><td style="padding:2px 12px;color:{TIERS['YELLOW'].color}">🟡 Yellow</td><td style="padding:2px 12px">≥ {watch:.0f}</td></tr>
  <tr><td style="padding:2px 12px;color:{TIERS['ORANGE'].color}">🟠 Orange</td><td style="padding:2px 12px">≥ {warning:.0f}</td></tr>
  <tr><td style="padding:2px 12px;color:{TIERS['RED'].color}">🔴 Red</td><td style="padding:2px 12px">≥ {curtail:.0f}</td></tr>
</table>

<h3 style="margin-bottom:4px">Hourly forecast (hour beginning ET, GW)</h3>
<table style="border-collapse:collapse;border:1px solid #DDD">
  <tr style="background:#1F3864;color:#FFF">
    <th style="padding:4px 10px;text-align:left">Hour</th>
    <th style="padding:4px 10px;text-align:right">GW</th>
  </tr>
  {rows_html}
</table>

<p style="color:#888;font-size:11px;margin-top:24px">
  Source: PJM Data Miner 2 feed <code>{PJM_FORECAST_FEED}</code>.
  Your capacity charge on the Dynegy line scales with Peak Load Contribution (PLC),
  set by your average demand during the five highest PJM RTO peak hours from
  June 1 – September 30. Curtailing during those five hours is the only way to lower it.
</p>
</body></html>
"""
    return subject, plain, html


def send_email(subject: str, plain: str, html: str, *, from_addr: str, to_addrs: list[str],
               smtp_host: str, smtp_port: int, smtp_user: str, smtp_pass: str,
               reply_to: str | None = None) -> None:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs)
    if reply_to:
        msg["Reply-To"] = reply_to
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls(context=context)
        smtp.ehlo()
        smtp.login(smtp_user, smtp_pass)
        smtp.send_message(msg)


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------

def parse_thresholds(env_value: str | None) -> tuple[int, int, int]:
    if not env_value:
        return DEFAULT_THRESHOLDS_MW
    parts = [int(x.strip()) for x in env_value.split(",")]
    if len(parts) != 3:
        raise SystemExit("ALERT_THRESHOLDS must be 'watch,warning,curtail' in MW.")
    return tuple(parts)  # type: ignore[return-value]


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="PJM 5CP Peak Load Alert")
    p.add_argument("--dry-run", action="store_true", help="Print email body instead of sending.")
    p.add_argument("--csv", help="Use a local PJM forecast CSV instead of the live API.")
    p.add_argument("--date", help="Target date YYYY-MM-DD (default: today, ET).")
    args = p.parse_args(argv)

    target_date = (
        datetime.strptime(args.date, "%Y-%m-%d").date()
        if args.date
        else datetime.now(EASTERN).date()
    )
    thresholds = parse_thresholds(os.environ.get("ALERT_THRESHOLDS"))

    if args.csv:
        rows = fetch_forecast_csv(args.csv, target_date)
    else:
        api_key = os.environ.get("PJM_API_KEY")
        if not api_key:
            print("ERROR: PJM_API_KEY environment variable not set.", file=sys.stderr)
            print("       Either set it, or pass --csv PATH for offline testing.", file=sys.stderr)
            return 2
        rows = fetch_forecast_live(api_key, target_date)
        # Normalize live rows into the internal shape
        for r in rows:
            r["_hour_ept"] = parse_pjm_dt(r["forecast_hour_beginning_ept"])
            r["_evaluated_ept"] = parse_pjm_dt(r["evaluated_at_ept"])
            r["_load_mw"] = float(r["forecast_load_mw"])

    dp = compute_day_peak(rows, target_date, thresholds)
    if dp is None:
        print(f"No PJM RTO forecast rows found for {target_date}.", file=sys.stderr)
        return 1

    subject, plain, html = render_email(dp, thresholds)

    if args.dry_run:
        print("=" * 70)
        print("SUBJECT:", subject)
        print("=" * 70)
        print(plain)
        return 0

    to_raw = os.environ.get("MAIL_TO", "")
    to_addrs = [x.strip() for x in to_raw.split(",") if x.strip()]
    if not to_addrs:
        print("ERROR: MAIL_TO not set (comma-separated list).", file=sys.stderr)
        return 2

    send_email(
        subject, plain, html,
        from_addr=os.environ.get("MAIL_FROM", os.environ["SMTP_USER"]),
        to_addrs=to_addrs,
        smtp_host=os.environ.get("SMTP_HOST", "smtp.office365.com"),
        smtp_port=int(os.environ.get("SMTP_PORT", "587")),
        smtp_user=os.environ["SMTP_USER"],
        smtp_pass=os.environ["SMTP_PASS"],
        reply_to=os.environ.get("MAIL_REPLY_TO"),
    )
    print(f"Sent {dp.tier.name} alert to {len(to_addrs)} recipient(s) "
          f"— peak {dp.peak_mw/1000:,.1f} GW at {dp.peak_hour_ept:%H:%M ET}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
