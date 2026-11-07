"""
Laptop Black Friday price tracker.

Price history lives in data/history.csv (source of truth). Every run rebuilds
the Excel workbook and an HTML report from it, rates each laptop's best price
(Excellent / Good / Fair / Poor), gives a BUY NOW / WAIT recommendation and can
email the report via Gmail SMTP.

Commands
  python tracker.py add <prices.json>     add a day's observations
  python tracker.py build                 rebuild Excel + HTML report, print summary
  python tracker.py email                 email the latest report
  python tracker.py daily <prices.json>   add + build + email (what the scheduled task runs)

prices.json format:
{
  "date": "2026-11-08",
  "observations": [
    {"laptop_id": "omen16", "store": "Amazon.de", "price": 1399.0, "currency": "EUR",
     "url": "https://...", "in_stock": true, "sku": "16-ap1190ng", "note": "32GB variant"}
  ]
}
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import smtplib
import sys
import urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
REPORTS = ROOT / "reports"
HISTORY = DATA / "history.csv"
FIELDS = ["date", "laptop_id", "store", "price", "currency", "landed_eur", "url",
          "in_stock", "sku", "note"]

RATING_ORDER = {"Excellent": 0, "Good": 1, "Fair": 2, "Poor": 3, "No data": 4}


# ---------------------------------------------------------------- loading
def load_json(name: str) -> dict:
    return json.loads((ROOT / name).read_text(encoding="utf-8"))


def cfg() -> dict:
    return load_json("config.json")


def laptops() -> list[dict]:
    return load_json("laptops.json")["laptops"]


def read_history() -> list[dict]:
    if not HISTORY.exists():
        return []
    with HISTORY.open(newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["price"] = float(r["price"])
        r["landed_eur"] = float(r["landed_eur"])
    return rows


# ---------------------------------------------------------------- money
_fx_cache: dict[str, float] = {}


def fx_to_eur(currency: str) -> float:
    """EUR per 1 unit of currency (ECB rate via frankfurter.app, config fallback)."""
    currency = currency.upper()
    if currency == "EUR":
        return 1.0
    if currency in _fx_cache:
        return _fx_cache[currency]
    rate = None
    try:
        url = f"https://api.frankfurter.app/latest?from={currency}&to=EUR"
        with urllib.request.urlopen(url, timeout=15) as r:
            rate = float(json.load(r)["rates"]["EUR"])
    except Exception:
        rate = float(cfg().get("fx_fallback", {}).get(currency, 0)) or None
    if rate is None:
        raise ValueError(f"No FX rate for {currency}")
    _fx_cache[currency] = rate
    return rate


def landed_eur(store: str, price: float, currency: str) -> float:
    stores = cfg()["stores"]
    s = stores.get(store) or stores["Other"]
    eur = price * fx_to_eur(currency)
    return round(eur * float(s["vat_factor"]) + float(s["shipping_eur"]), 2)


# ---------------------------------------------------------------- add
def add(path: str) -> int:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    day = payload.get("date") or dt.date.today().isoformat()
    known = {l["id"] for l in laptops()}
    stores = cfg()["stores"]
    existing = read_history()
    # replace any earlier rows for the same day/laptop/store (re-runs are idempotent)
    new_keys = set()
    new_rows = []
    for o in payload["observations"]:
        if o["laptop_id"] not in known:
            print(f"  ! unknown laptop_id {o['laptop_id']!r} - skipped")
            continue
        if o.get("price") in (None, "", 0):
            continue
        store = o["store"] if o["store"] in stores else o["store"]
        cur = o.get("currency", "EUR")
        price = float(o["price"])
        new_rows.append({
            "date": day, "laptop_id": o["laptop_id"], "store": store,
            "price": price, "currency": cur,
            "landed_eur": landed_eur(store, price, cur),
            "url": o.get("url", ""), "in_stock": str(o.get("in_stock", True)).lower(),
            "sku": o.get("sku", ""), "note": o.get("note", ""),
        })
        new_keys.add((day, o["laptop_id"], store))
    kept = [r for r in existing if (r["date"], r["laptop_id"], r["store"]) not in new_keys]
    DATA.mkdir(exist_ok=True)
    with HISTORY.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for r in sorted(kept + new_rows, key=lambda r: (r["date"], r["laptop_id"], r["landed_eur"])):
            w.writerow(r)
    print(f"Added {len(new_rows)} observations for {day}.")
    return len(new_rows)


# ---------------------------------------------------------------- analysis
def phase(today: dt.date) -> str:
    sd = cfg()["sale_dates"]
    bf = dt.date.fromisoformat(sd["black_friday"])
    cm = dt.date.fromisoformat(sd["cyber_monday"])
    if today < bf - dt.timedelta(days=7):
        return "pre"          # early deals - BF usually still beats them
    if today <= cm + dt.timedelta(days=1):
        return "bf"           # Black Friday week -> Cyber Monday
    return "post"


def rate(price: float | None, lap: dict) -> str:
    if price is None:
        return "No data"
    if price <= lap["excellent"]:
        return "Excellent"
    if price <= lap["good"]:
        return "Good"
    if price <= lap["fair"]:
        return "Fair"
    return "Poor"


def recommend(rating: str, ph: str, at_low: bool, budget_max: float, price: float | None) -> str:
    if rating == "No data":
        return "NO DATA - check manually"
    over = price is not None and price > budget_max
    if rating == "Excellent":
        return "BUY NOW" + (" (over budget!)" if over else "")
    if rating == "Good":
        if ph == "pre":
            return "BUY if lowest seen, else WAIT for BF" if not at_low else "BUY - lowest price tracked"
        return "BUY - good BF price" + (" (over budget)" if over else "")
    if rating == "Fair":
        if ph in ("pre", "bf"):
            return "WAIT - Black Friday should beat this"
        return "FAIR - OK if you need it now, else wait for Jan sales"
    return "SKIP / WAIT - overpriced"


def analyse(history: list[dict]) -> tuple[list[dict], str]:
    today = dt.date.today()
    ph = phase(today)
    dates = sorted({r["date"] for r in history})
    latest = dates[-1] if dates else None
    results = []
    for lap in laptops():
        rows = [r for r in history if r["laptop_id"] == lap["id"]
                and r["in_stock"] != "false"]
        by_date: dict[str, dict] = {}
        for r in rows:
            if r["date"] not in by_date or r["landed_eur"] < by_date[r["date"]]["landed_eur"]:
                by_date[r["date"]] = r
        lap_dates = sorted(by_date)
        cur = by_date[lap_dates[-1]] if lap_dates else None
        prev = by_date[lap_dates[-2]] if len(lap_dates) > 1 else None
        all_low = min((r["landed_eur"] for r in by_date.values()), default=None)
        price = cur["landed_eur"] if cur else None
        at_low = price is not None and all_low is not None and price <= all_low + 0.01 and len(lap_dates) >= 3
        stale = cur is not None and latest is not None and cur["date"] != latest
        rating = rate(price, lap)
        rec = recommend(rating, ph, at_low, 1600, price)
        change = None
        if cur and prev:
            change = round(price - prev["landed_eur"], 2)
        # value = WoW score per EUR1000
        value = round(lap["wow_score"] / price * 1000, 2) if price else None
        results.append({
            "lap": lap, "cur": cur, "prev": prev, "price": price, "change": change,
            "all_low": all_low, "at_low": at_low, "stale": stale, "rating": rating,
            "rec": rec, "value": value, "days": len(lap_dates),
            "daily": {d: by_date[d]["landed_eur"] for d in lap_dates},
        })
    results.sort(key=lambda x: (RATING_ORDER[x["rating"]], -(x["value"] or 0)))
    return results, ph


def top_pick(results: list[dict]) -> dict | None:
    cands = [r for r in results if r["price"] and r["rating"] in ("Excellent", "Good")
             and r["price"] <= 1600]
    if not cands:
        cands = [r for r in results if r["price"] and r["price"] <= 1600]
    return max(cands, key=lambda r: r["value"], default=None)


def headline(results, ph) -> str:
    buys = [r for r in results if r["rec"].startswith("BUY NOW")]
    tp = top_pick(results)
    days_bf = (dt.date.fromisoformat(cfg()["sale_dates"]["black_friday"]) - dt.date.today()).days
    if buys:
        b = buys[0]
        return f"BUY NOW: {b['lap']['name']} EUR{b['price']:,.0f} ({b['rating']}) at {b['cur']['store']}"
    if tp:
        when = f"{days_bf} days to Black Friday" if days_bf > 0 else "Black Friday period"
        return f"No must-buy yet - best value {tp['lap']['name']} EUR{tp['price']:,.0f} ({tp['rating']}). {when}."
    return "No price data yet."


# ---------------------------------------------------------------- excel
def build_excel(results, ph, history) -> Path:
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    c = cfg()
    wb = Workbook()
    hdr_fill = PatternFill("solid", fgColor="1F2937")
    hdr_font = Font(color="FFFFFF", bold=True)
    fills = {
        "Excellent": PatternFill("solid", fgColor="16A34A"),
        "Good": PatternFill("solid", fgColor="86EFAC"),
        "Fair": PatternFill("solid", fgColor="FDE68A"),
        "Poor": PatternFill("solid", fgColor="FCA5A5"),
        "No data": PatternFill("solid", fgColor="E5E7EB"),
    }

    def header(ws, row, cols):
        for i, h in enumerate(cols, 1):
            cell = ws.cell(row=row, column=i, value=h)
            cell.fill, cell.font = hdr_fill, hdr_font
            cell.alignment = Alignment(wrap_text=True, vertical="center")

    # --- Dashboard
    ws = wb.active
    ws.title = "Dashboard"
    ws["A1"] = f"Laptop Black Friday Tracker - updated {dt.datetime.now():%Y-%m-%d %H:%M}"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"] = headline(results, ph)
    ws["A2"].font = Font(size=12, bold=True, color="B91C1C")
    ws["A3"] = ("Prices = cheapest in-stock offer, 'landed' in Ireland (IE VAT, FX, shipping - see Settings). "
                "Budget EUR1,200-1,600. Phase: " + {"pre": "pre-Black-Friday", "bf": "Black Friday week",
                                                    "post": "after Cyber Monday"}[ph])
    cols = ["Laptop", "CPU", "GPU", "WoW score", "Best price EUR (landed)", "Store", "Shop price",
            "Change vs prev", "Lowest seen EUR", "Fair EUR", "Good EUR", "Excellent EUR",
            "Deal rating", "Recommendation", "Value (score/EUR1k)", "Days tracked", "Last seen", "Link", "Notes"]
    header(ws, 5, cols)
    for i, r in enumerate(results, 6):
        lap, cur = r["lap"], r["cur"]
        vals = [lap["name"], lap["cpu"], lap["gpu"], lap["wow_score"], r["price"],
                cur["store"] if cur else "", f"{cur['price']:,.2f} {cur['currency']}" if cur else "",
                r["change"], r["all_low"], lap["fair"], lap["good"], lap["excellent"],
                r["rating"], r["rec"], r["value"], r["days"], cur["date"] if cur else "",
                cur["url"] if cur else "", lap["notes"]]
        for j, v in enumerate(vals, 1):
            ws.cell(row=i, column=j, value=v)
        ws.cell(row=i, column=13).fill = fills[r["rating"]]
        if r["rec"].startswith("BUY NOW"):
            ws.cell(row=i, column=14).font = Font(bold=True, color="15803D")
        if cur and cur["url"]:
            ws.cell(row=i, column=18).hyperlink = cur["url"]
            ws.cell(row=i, column=18).value = "open"
            ws.cell(row=i, column=18).font = Font(color="2563EB", underline="single")
        for col in (5, 9, 10, 11, 12):
            ws.cell(row=i, column=col).number_format = '€#,##0'
        ws.cell(row=i, column=8).number_format = '+€#,##0;-€#,##0;0'
    last = 5 + len(results)
    ws.conditional_formatting.add(f"H6:H{last}", CellIsRule(operator="lessThan", formula=["0"],
                                  font=Font(color="15803D", bold=True)))
    ws.conditional_formatting.add(f"H6:H{last}", CellIsRule(operator="greaterThan", formula=["0"],
                                  font=Font(color="B91C1C", bold=True)))
    widths = [34, 22, 12, 9, 13, 22, 16, 11, 12, 9, 9, 10, 11, 40, 11, 8, 11, 7, 60]
    for i, wdt in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = wdt
    ws.freeze_panes = "B6"

    # --- Daily best (pivot + chart)
    wd = wb.create_sheet("Daily Best")
    dates = sorted({r["date"] for r in history})
    header(wd, 1, ["Date"] + [r["lap"]["name"] for r in results])
    for i, d in enumerate(dates, 2):
        wd.cell(row=i, column=1, value=d)
        for j, r in enumerate(results, 2):
            v = r["daily"].get(d)
            if v is not None:
                wd.cell(row=i, column=j, value=v).number_format = '€#,##0'
    wd.column_dimensions["A"].width = 12
    for j in range(2, len(results) + 2):
        wd.column_dimensions[get_column_letter(j)].width = 16
    if len(dates) >= 1:
        ch = LineChart()
        ch.title = "Best landed price per laptop (EUR)"
        ch.y_axis.title = "EUR"
        ch.height, ch.width = 12, 28
        data = Reference(wd, min_col=2, max_col=len(results) + 1, min_row=1, max_row=len(dates) + 1)
        ch.add_data(data, titles_from_data=True)
        ch.set_categories(Reference(wd, min_col=1, min_row=2, max_row=len(dates) + 1))
        wd.add_chart(ch, f"A{len(dates) + 4}")

    # --- Price history
    wh = wb.create_sheet("Price History")
    names = {l["id"]: l["name"] for l in laptops()}
    hcols = ["Date", "Laptop", "Store", "Shop price", "Currency", "Landed EUR", "In stock", "SKU", "Note", "URL"]
    header(wh, 1, hcols)
    for i, r in enumerate(sorted(history, key=lambda r: (r["date"], r["laptop_id"], r["landed_eur"]), reverse=True), 2):
        vals = [r["date"], names.get(r["laptop_id"], r["laptop_id"]), r["store"], r["price"], r["currency"],
                r["landed_eur"], r["in_stock"], r["sku"], r["note"], r["url"]]
        for j, v in enumerate(vals, 1):
            wh.cell(row=i, column=j, value=v)
        wh.cell(row=i, column=6).number_format = '€#,##0.00'
    for i, wdt in enumerate([11, 34, 22, 11, 8, 11, 8, 22, 40, 60], 1):
        wh.column_dimensions[get_column_letter(i)].width = wdt
    wh.auto_filter.ref = f"A1:J{max(2, len(history) + 1)}"
    wh.freeze_panes = "A2"

    # --- Watchlist
    wl = wb.create_sheet("Watchlist")
    header(wl, 1, ["ID", "Laptop", "CPU", "GPU", "RAM", "Screen", "WoW score", "Fair", "Good", "Excellent", "Notes"])
    for i, l in enumerate(laptops(), 2):
        for j, v in enumerate([l["id"], l["name"], l["cpu"], l["gpu"], l["ram"], l["screen"], l["wow_score"],
                               l["fair"], l["good"], l["excellent"], l["notes"]], 1):
            wl.cell(row=i, column=j, value=v)
    wl.cell(row=len(laptops()) + 3, column=1, value="Edit thresholds/models in laptops.json (this sheet is regenerated daily).")
    for i, wdt in enumerate([11, 34, 26, 12, 9, 34, 9, 8, 8, 9, 70], 1):
        wl.column_dimensions[get_column_letter(i)].width = wdt

    # --- Settings / stores
    st = wb.create_sheet("Settings")
    header(st, 1, ["Store", "Country", "Currency", "VAT factor -> IE", "Shipping EUR", "Ships to IE?"])
    for i, (name, s) in enumerate(c["stores"].items(), 2):
        for j, v in enumerate([name, s["country"], s["currency"], s["vat_factor"], s["shipping_eur"], s["ships_to_ie"]], 1):
            st.cell(row=i, column=j, value=v)
    n = len(c["stores"]) + 3
    st.cell(row=n, column=1, value=f"GBP->EUR used today: {fx_to_eur('GBP'):.4f} (ECB)")
    st.cell(row=n + 1, column=1, value="Rating: <= Excellent price = Excellent, <= Good = Good, <= Fair = Fair, else Poor.")
    st.cell(row=n + 2, column=1, value="Edit config.json to change stores/VAT/shipping, laptops.json for models and price targets.")
    for i, wdt in enumerate([26, 9, 9, 15, 12, 34], 1):
        st.column_dimensions[get_column_letter(i)].width = wdt

    out = ROOT / c["excel_file"]
    try:
        wb.save(out)
    except PermissionError:  # file open in Excel
        out = ROOT / (out.stem + f"_{dt.date.today():%Y%m%d}.xlsx")
        wb.save(out)
        print(f"  ! Excel file was open - saved to {out.name} instead")
    return out


# ---------------------------------------------------------------- html report
def build_html(results, ph) -> str:
    colors = {"Excellent": "#16a34a", "Good": "#65a30d", "Fair": "#ca8a04", "Poor": "#dc2626", "No data": "#6b7280"}
    tp = top_pick(results)
    rows = []
    for r in results:
        lap, cur = r["lap"], r["cur"]
        price = f"€{r['price']:,.0f}" if r["price"] else "-"
        chg = ""
        if r["change"]:
            col = "#16a34a" if r["change"] < 0 else "#dc2626"
            chg = f"<span style='color:{col}'>{r['change']:+,.0f}</span>"
        store = escape(cur["store"]) if cur else "-"
        if cur and cur["url"]:
            store = f"<a href='{escape(cur['url'])}'>{store}</a>"
        low = f"€{r['all_low']:,.0f}" if r["all_low"] else "-"
        rows.append(
            f"<tr><td><b>{escape(lap['name'])}</b><br><small>{escape(lap['cpu'])} · {escape(lap['gpu'])} · WoW {lap['wow_score']}/10</small></td>"
            f"<td style='text-align:right'><b>{price}</b> {chg}</td><td>{store}</td>"
            f"<td style='text-align:right'>{low}</td>"
            f"<td style='text-align:right'><small>€{lap['excellent']:,} / €{lap['good']:,} / €{lap['fair']:,}</small></td>"
            f"<td><span style='background:{colors[r['rating']]};color:#fff;padding:2px 6px;border-radius:4px'>{r['rating']}</span></td>"
            f"<td>{escape(r['rec'])}{' <small>(stale)</small>' if r['stale'] else ''}</td></tr>")
    tp_line = (f"<p><b>Best value in budget:</b> {escape(tp['lap']['name'])} – €{tp['price']:,.0f} at "
               f"{escape(tp['cur']['store'])} ({tp['rating']})</p>") if tp else ""
    phase_txt = {"pre": "Pre-Black-Friday: early deals are rarely the lowest – only jump on Excellent ones.",
                 "bf": "Black Friday week: this is when to buy – Good or Excellent = go.",
                 "post": "After Cyber Monday: prices tend to creep back up (RAM shortage) – buy Good deals."}[ph]
    return f"""<html><head><meta charset="utf-8"></head><body style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111">
<h2 style="margin-bottom:4px">Laptop tracker – {dt.date.today():%a %d %b %Y}</h2>
<p style="font-size:16px"><b>{escape(headline(results, ph))}</b></p>
{tp_line}<p style="color:#555">{phase_txt}</p>
<table cellpadding="6" style="border-collapse:collapse;border:1px solid #ddd">
<tr style="background:#1f2937;color:#fff"><th>Laptop</th><th>Best landed €</th><th>Store</th><th>Lowest seen</th>
<th>Excellent / Good / Fair</th><th>Rating</th><th>Advice</th></tr>
{''.join(rows)}</table>
<p style="color:#777;font-size:12px">Landed = delivered to Ireland incl. 23% VAT, FX and shipping estimate.
Check the exact SKU (CPU/RAM/screen) before buying. Full history in Laptop_Price_Tracker.xlsx.</p>
</body></html>"""


def build() -> tuple[list[dict], str, str]:
    history = read_history()
    results, ph = analyse(history)
    xlsx = build_excel(results, ph, history)
    html = build_html(results, ph)
    REPORTS.mkdir(exist_ok=True)
    rp = REPORTS / f"report_{dt.date.today():%Y-%m-%d}.html"
    rp.write_text(html, encoding="utf-8")
    (REPORTS / "latest.html").write_text(html, encoding="utf-8")
    print(headline(results, ph))
    for r in results:
        p = f"EUR{r['price']:,.0f}" if r["price"] else "   -  "
        print(f"  {r['rating']:9} {p:>9}  {r['lap']['name'][:40]:40} {r['rec']}")
    print(f"Excel: {xlsx}\nReport: {rp}")
    return results, ph, html


# ---------------------------------------------------------------- email
def email(results=None, ph=None, html=None) -> None:
    c = cfg()["email"]
    if not c.get("enabled"):
        print("Email disabled in config.json")
        return
    if html is None:
        results, ph, html = build()
    import os
    pw = os.environ.get("GMAIL_APP_PASSWORD", "")  # cloud: set as environment variable
    sec_path = ROOT / "secrets.json"
    if not pw and sec_path.exists():
        pw = json.loads(sec_path.read_text(encoding="utf-8")).get("gmail_app_password", "")
    pw = pw.replace(" ", "")
    if not pw or pw.startswith("xxxx") or pw.upper().startswith("PASTE"):
        print("! No Gmail app password (GMAIL_APP_PASSWORD env var or secrets.json). Email not sent.")
        return
    if not pw or pw.startswith("xxxx") or pw.upper().startswith("PASTE"):
        print("! No Gmail app password in secrets.json. Email not sent.")
        return
    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"[Laptop tracker] {headline(results, ph)}"
    msg["From"], msg["To"] = c["from"], c["to"]
    msg.attach(MIMEText(html, "html", "utf-8"))
    xlsx = ROOT / cfg()["excel_file"]
    if xlsx.exists():  # attach the workbook so it can be opened on a phone
        from email.mime.application import MIMEApplication
        part = MIMEApplication(xlsx.read_bytes(), Name=xlsx.name)
        part["Content-Disposition"] = f'attachment; filename="{xlsx.name}"'
        msg.attach(part)
    with smtplib.SMTP(c["smtp_host"], c["smtp_port"], timeout=30) as s:
        s.starttls()
        s.login(c["from"], pw)
        s.send_message(msg)
    print(f"Email sent to {c['to']}")


# ---------------------------------------------------------------- cli
def main(argv: list[str]) -> None:
    if len(argv) < 2:
        print(__doc__)
        return
    cmd = argv[1]
    if cmd == "add":
        add(argv[2])
    elif cmd == "build":
        build()
    elif cmd == "email":
        email()
    elif cmd == "daily":
        add(argv[2])
        results, ph, html = build()
        email(results, ph, html)
    else:
        print(__doc__)


if __name__ == "__main__":
    main(sys.argv)
