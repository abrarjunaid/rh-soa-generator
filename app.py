"""
Radiant Homes — SOA Generator Web App
======================================
Upload a P&L workbook, select month, generate all owner statement PDFs.

Run locally:  python app.py
Deploy:       Push to GitHub → connect to Railway.app
"""

import base64
import calendar
import io
import os
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from playwright.sync_api import sync_playwright

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max upload

# ─── Helpers ───────────────────────────────────────────────────────────────────

def month_key(dt):
    if not isinstance(dt, datetime):
        return None
    return f"{dt.year}-{str(dt.month).zfill(2)}"


def month_label(key):
    months = ["", "January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]
    y, m = key.split("-")
    return f"{months[int(m)]} {y}"


def month_short(key):
    months = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    y, m = key.split("-")
    return f"{months[int(m)]} '{y[2:]}"


def days_in_month(key):
    y, m = key.split("-")
    return calendar.monthrange(int(y), int(m))[1]


def next_month_label(key):
    months = ["", "January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]
    y, m = int(key.split("-")[0]), int(key.split("-")[1])
    nm = m + 1 if m < 12 else 1
    ny = y if m < 12 else y + 1
    return f"{months[nm]} {ny}"


def fmt(n, decimals=2):
    return f"{n:,.{decimals}f}"


def format_date(dt):
    if not isinstance(dt, datetime):
        return ""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{dt.day} {months[dt.month - 1]}"


# ─── Data Extraction ──────────────────────────────────────────────────────────

def load_unit_registry(wb):
    ws = wb["Unit Registry"]
    units = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=False):
        code = row[1].value
        building = row[2].value
        model = row[3].value
        if not code or not building or str(model).strip() != "Revenue Share":
            continue
        if str(code).strip() == "Unit Code":
            continue
        units.append({
            "code": str(code).strip(),
            "building": str(building).strip(),
            "owner": str(row[6].value or "[Owner Name]").strip(),
            "email": str(row[7].value or "[Email]").strip(),
            "phone": str(row[8].value or "[Phone]").strip(),
            "active": str(row[10].value or "").strip(),
        })
    return units


def get_available_months(wb, units):
    months_set = set()
    for unit in units:
        try:
            ws = wb[unit["code"]]
        except KeyError:
            continue
        for col in range(2, ws.max_column + 1):
            cell_val = ws.cell(row=3, column=col).value
            if isinstance(cell_val, datetime):
                months_set.add(month_key(cell_val))
    return sorted(months_set, reverse=True)


def load_pnl(wb, unit_code, month):
    try:
        ws = wb[unit_code]
    except KeyError:
        return None

    col_idx = None
    for col in range(2, ws.max_column + 1):
        cell_val = ws.cell(row=3, column=col).value
        if isinstance(cell_val, datetime) and month_key(cell_val) == month:
            col_idx = col
            break

    if col_idx is None:
        return None

    def v(row_num):
        val = ws.cell(row=row_num, column=col_idx).value
        return round(float(val), 2) if isinstance(val, (int, float)) else 0.0

    pnl = {
        "total_gross": v(15), "platform_fees": v(16), "payment_charges": v(17),
        "net_earned": v(19), "cleaning_retained": v(21), "tourism_retained": v(22),
        "rev_net_retained": v(23), "total_owner_expenses": v(28),
        "net_before_mgmt": v(29), "mgmt_fee": v(30), "owner_payout": v(31),
    }

    if pnl["owner_payout"] == 0 and pnl["total_gross"] == 0:
        return None
    return pnl


def load_bookings(wb, unit_code, month):
    ws = wb["Sales"]
    header_row = None
    for r in range(1, 6):
        val = ws.cell(row=r, column=1).value
        if val and "Hostaway" in str(val):
            header_row = r
            break
    if header_row is None:
        return []

    bookings = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=False):
        prop = row[3].value
        sale_month = row[7].value
        if (str(prop).strip() == unit_code and
                isinstance(sale_month, datetime) and
                month_key(sale_month) == month):
            bookings.append({
                "guest": str(row[2].value or ""),
                "platform": str(row[6].value or ""),
                "checkin": row[8].value,
                "checkout": row[9].value,
                "nights": int(row[10].value or 0),
                "cleaning": round(float(row[13].value or 0), 2),
                "tourism": round(float(row[15].value or 0), 2),
                "guest_paid": round(float(row[22].value or 0), 2),
                "host_fee_total": round(float(row[25].value or 0), 2),
                "payment_charges": round(float(row[26].value or 0), 2),
                "remitted": round(float(row[31].value or 0), 2),
            })
    return bookings


# ─── SOA Calculation ──────────────────────────────────────────────────────────

def calculate_soa(unit, pnl, bookings, month):
    running_pm = 0.0
    rows = []

    for i, b in enumerate(bookings):
        rev = b["guest_paid"]
        cleaning = b["cleaning"]
        tourism = b["tourism"]
        commission = round(abs(b["host_fee_total"]) + abs(b["payment_charges"]), 2)
        net = b["remitted"]
        rev_net_ret = net - cleaning - tourism
        pm = round(rev_net_ret * 0.15, 2)
        running_pm += pm
        gross = round(rev_net_ret - pm, 2)

        ch_class, ch_label = "ch-direct", "Direct"
        plat = b["platform"].lower()
        if "airbnb" in plat:
            ch_class, ch_label = "ch-airbnb", "Airbnb"
        elif "booking" in plat:
            ch_class, ch_label = "ch-booking", "Booking"

        rows.append({
            "num": i + 1, "guest": b["guest"], "ch_class": ch_class, "ch_label": ch_label,
            "checkin": format_date(b["checkin"]), "checkout": format_date(b["checkout"]),
            "nights": b["nights"], "rev": rev, "cleaning": cleaning, "commission": commission,
            "net": net, "pm": pm, "gross": gross,
        })

    # Adjust PM to match PnL
    pm_target = abs(pnl["mgmt_fee"])
    pm_diff = round(running_pm - pm_target, 2)
    if abs(pm_diff) > 0.001 and rows:
        largest = max(rows, key=lambda r: r["net"])
        largest["pm"] = round(largest["pm"] - pm_diff, 2)
        largest["gross"] = round(largest["gross"] + pm_diff, 2)

    totals = {k: round(sum(r[k] for r in rows), 2) for k in ["rev", "cleaning", "commission", "net", "pm", "gross"]}
    totals["nights"] = sum(r["nights"] for r in rows)

    available = days_in_month(month)
    expenses = abs(pnl["total_owner_expenses"])
    fees_received = round(abs(pnl["cleaning_retained"]) + abs(pnl["tourism_retained"]), 2)
    total_ded = round(fees_received + expenses + abs(pnl["mgmt_fee"]) + abs(pnl["platform_fees"]) + abs(pnl["payment_charges"]), 2)

    unit_number = unit["code"].split(" ")[1] if " " in unit["code"] else unit["code"]

    return {
        "unit": unit, "month": month,
        "property_name": f"{unit['building']} {unit_number}",
        "rows": rows, "totals": totals, "available": available,
        "expenses": expenses, "fees_received": fees_received,
        "deductions": {
            "fees_received": fees_received, "utilities": expenses,
            "mgmt_fee": abs(pnl["mgmt_fee"]), "platform_fees": abs(pnl["platform_fees"]),
            "payment_charges": abs(pnl["payment_charges"]), "total": total_ded,
        },
        "kpi": {
            "owner_gross": round(totals["gross"]),
            "occupancy": round((totals["nights"] / available) * 100),
            "booked": totals["nights"], "available": available,
            "reservations": len(rows),
            "net_payout": round(pnl["owner_payout"]),
            "net_payout_exact": round(pnl["owner_payout"], 2),
        },
    }


# ─── HTML Template ────────────────────────────────────────────────────────────

def generate_html(soa, logo_b64=None):
    u = soa["unit"]
    k = soa["kpi"]
    t = soa["totals"]
    d = soa["deductions"]
    m = soa["month"]
    y, mn = m.split("-")
    ms = ["", "January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
    period = f"{ms[int(mn)]} 1 — {days_in_month(m)}, {y}"

    logo = f'<img src="data:image/png;base64,{logo_b64}" style="height:36px;width:auto">' if logo_b64 else '<div style="font-size:24px;font-weight:800;color:#1a1d24">Radiant h<span style="color:#1565a0">o</span>mes.</div>'

    brows = ""
    for r in soa["rows"]:
        dim = ' style="color:#6b7280;opacity:0.45"' if r["rev"] == 0 else ""
        brows += f'''<tr><td{dim}>{r["num"]}</td><td{dim}>{r["guest"]}</td>
        <td><span class="ch {r["ch_class"]}">{r["ch_label"]}</span></td>
        <td{dim}>{r["checkin"]}</td><td{dim}>{r["checkout"]}</td>
        <td class="r"{dim}>{r["nights"]}</td><td class="r"{dim}>{fmt(r["rev"])}</td><td class="r"{dim}>{fmt(r["cleaning"])}</td>
        <td class="r"{dim}>{fmt(r["commission"])}</td><td class="r"{dim}>{fmt(r["net"])}</td><td class="r"{dim}>{fmt(r["pm"])}</td><td class="r"{dim}>{fmt(r["gross"])}</td></tr>'''

    return f'''<!DOCTYPE html><html><head><meta charset="UTF-8">
<link href="https://fonts.bunny.net/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Outfit',sans-serif;background:#fff;color:#1a1d24;-webkit-font-smoothing:antialiased}}
.page{{width:210mm;margin:0 auto;position:relative;overflow:hidden}}
.page::before{{content:'';position:absolute;top:0;left:0;right:0;height:5px;background:linear-gradient(90deg,#0d4a75,#1565a0,#0d4a75)}}
table{{width:100%;border-collapse:collapse;font-size:11px}}
thead th{{text-align:left;padding:10px 6px;font-size:8.5px;letter-spacing:1.2px;text-transform:uppercase;color:#6b7280;background:#f7f9fb;border-bottom:2px solid #1565a0;font-weight:600;white-space:nowrap}}
thead th.r,tbody td.r,tfoot td.r{{text-align:right}}
tbody td{{padding:9px 6px;border-bottom:1px solid #eaeff4;white-space:nowrap;font-variant-numeric:tabular-nums}}
tfoot td{{padding:12px 6px;font-weight:700;border-top:2px solid #1565a0;background:#f7f9fb;font-size:11.5px}}
.ch{{display:inline-block;font-size:8px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;padding:3px 8px;border-radius:4px}}
.ch-airbnb{{background:rgba(217,79,79,.07);color:#d94f4f}}.ch-booking{{background:rgba(21,101,160,.07);color:#1565a0}}.ch-direct{{background:rgba(26,138,106,.07);color:#1a8a6a}}
</style></head><body>
<div class="page">
<div style="padding:32px 44px 0;display:flex;justify-content:space-between;align-items:center;position:relative;z-index:1">{logo}<div style="font-size:10px;font-weight:600;letter-spacing:2.5px;text-transform:uppercase;color:#1565a0;border:1.5px solid #1565a0;padding:6px 16px;border-radius:6px">Owner's Statement</div></div>
<div style="padding:26px 44px 22px;display:flex;justify-content:space-between;align-items:flex-end;border-bottom:1px solid #dde3ea"><div><div style="font-size:36px;font-weight:700;letter-spacing:-.5px;line-height:1.1">{soa["property_name"]}</div><div style="font-size:13px;color:#6b7280;margin-top:6px">{period}</div></div><div style="text-align:right"><div style="font-size:9px;letter-spacing:2.5px;text-transform:uppercase;color:#6b7280;margin-bottom:6px">Net Owner Payout</div><div style="font-size:42px;font-weight:800;color:#1565a0;letter-spacing:-1px"><span style="font-size:16px;font-weight:400;color:#6b7280">AED </span>{k["net_payout"]:,}</div></div></div>
<div style="padding:14px 44px;background:#f7f9fb;border-bottom:1px solid #eaeff4;display:flex;gap:40px;font-size:12px;color:#6b7280"><span><strong style="color:#1a1d24;font-weight:500">Owner:</strong> {u["owner"]}</span><span><strong style="color:#1a1d24;font-weight:500">Phone:</strong> {u["phone"]}</span><span><strong style="color:#1a1d24;font-weight:500">Email:</strong> {u["email"]}</span></div>
<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;padding:22px 44px">
<div style="border:1.5px solid #dde3ea;border-radius:12px;padding:20px;text-align:center;position:relative"><div style="position:absolute;top:0;left:28%;right:28%;height:3px;border-radius:0 0 3px 3px;background:#1565a0"></div><div style="font-size:26px;font-weight:700;color:#1565a0;margin-bottom:4px;letter-spacing:-.5px">{k["owner_gross"]:,}</div><div style="font-size:9px;letter-spacing:1.5px;text-transform:uppercase;color:#6b7280;font-weight:500">Owner Gross (AED)</div></div>
<div style="border:1.5px solid #dde3ea;border-radius:12px;padding:20px;text-align:center;position:relative"><div style="position:absolute;top:0;left:28%;right:28%;height:3px;border-radius:0 0 3px 3px;background:#1a8a6a"></div><div style="font-size:26px;font-weight:700;color:#1a8a6a;margin-bottom:4px;letter-spacing:-.5px">{k["occupancy"]}%</div><div style="font-size:9px;letter-spacing:1.5px;text-transform:uppercase;color:#6b7280;font-weight:500">Occupancy</div></div>
<div style="border:1.5px solid #dde3ea;border-radius:12px;padding:20px;text-align:center;position:relative"><div style="position:absolute;top:0;left:28%;right:28%;height:3px;border-radius:0 0 3px 3px;background:#c08b2e"></div><div style="font-size:26px;font-weight:700;color:#c08b2e;margin-bottom:4px;letter-spacing:-.5px">{k["booked"]} / {k["available"]}</div><div style="font-size:9px;letter-spacing:1.5px;text-transform:uppercase;color:#6b7280;font-weight:500">Booked / Available</div></div>
<div style="border:1.5px solid #dde3ea;border-radius:12px;padding:20px;text-align:center;position:relative"><div style="position:absolute;top:0;left:28%;right:28%;height:3px;border-radius:0 0 3px 3px;background:#d94f4f"></div><div style="font-size:26px;font-weight:700;color:#d94f4f;margin-bottom:4px;letter-spacing:-.5px">{k["reservations"]}</div><div style="font-size:9px;letter-spacing:1.5px;text-transform:uppercase;color:#6b7280;font-weight:500">Reservations</div></div>
</div>
<div style="font-size:9px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#1565a0;padding:14px 44px 10px;display:flex;align-items:center;gap:14px">Rental Activity Details — {ms[int(mn)]} {y}<span style="flex:1;height:1px;background:#dde3ea"></span></div>
<div style="padding:0 44px 12px"><table><thead><tr><th>#</th><th>Guest</th><th>Channel</th><th>In</th><th>Out</th><th class="r">Nts</th><th class="r">Booking Rev</th><th class="r">Cleaning</th><th class="r">Commission</th><th class="r">Net Rev</th><th class="r">PM 15%</th><th class="r">Gross</th></tr></thead><tbody>{brows}</tbody>
<tfoot><tr><td colspan="5">Total</td><td class="r">{t["nights"]}</td><td class="r">{fmt(t["rev"])}</td><td class="r">{fmt(t["cleaning"])}</td><td class="r">{fmt(t["commission"])}</td><td class="r">{fmt(t["net"])}</td><td class="r">{fmt(t["pm"])}</td><td class="r">{fmt(t["gross"])}</td></tr></tfoot></table></div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;padding:0 44px 18px">
<div style="border:1.5px solid #dde3ea;border-radius:12px;padding:22px"><div style="font-size:9px;letter-spacing:2.5px;text-transform:uppercase;color:#1565a0;font-weight:700;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #dde3ea">Expenses & Extras</div><div style="display:flex;justify-content:space-between;padding:6px 0;font-size:12px"><span style="color:#6b7280">Utilities & Service Charge — {month_short(m)}</span><span style="font-weight:500;font-variant-numeric:tabular-nums">{fmt(soa["expenses"])}</span></div><div style="display:flex;justify-content:space-between;padding:10px 0 6px;font-size:12px;font-weight:700;border-top:1.5px solid #1a1d24;margin-top:8px"><span>Total Expenses</span><span style="color:#d94f4f;font-variant-numeric:tabular-nums">AED {fmt(soa["expenses"])}</span></div></div>
<div style="border:1.5px solid #dde3ea;border-radius:12px;padding:22px"><div style="font-size:9px;letter-spacing:2.5px;text-transform:uppercase;color:#1565a0;font-weight:700;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #dde3ea">Deductions Breakdown</div>
<div style="display:flex;justify-content:space-between;padding:6px 0;font-size:12px"><span style="color:#6b7280">Fees Received (Cleaning + Tourism)</span><span style="font-weight:500;font-variant-numeric:tabular-nums">{fmt(d["fees_received"])}</span></div>
<div style="display:flex;justify-content:space-between;padding:6px 0;font-size:12px"><span style="color:#6b7280">Apartment Expenses (Utilities)</span><span style="font-weight:500;font-variant-numeric:tabular-nums">{fmt(d["utilities"])}</span></div>
<div style="display:flex;justify-content:space-between;padding:6px 0;font-size:12px"><span style="color:#6b7280">Management Fee (15%)</span><span style="font-weight:500;font-variant-numeric:tabular-nums">{fmt(d["mgmt_fee"])}</span></div>
<div style="display:flex;justify-content:space-between;padding:6px 0;font-size:12px"><span style="color:#6b7280">Platform Host Fees</span><span style="font-weight:500;font-variant-numeric:tabular-nums">{fmt(d["platform_fees"])}</span></div>
<div style="display:flex;justify-content:space-between;padding:6px 0;font-size:12px"><span style="color:#6b7280">Payment Charges</span><span style="font-weight:500;font-variant-numeric:tabular-nums">{fmt(d["payment_charges"])}</span></div>
<div style="display:flex;justify-content:space-between;padding:10px 0 6px;font-size:12px;font-weight:700;border-top:1.5px solid #1a1d24;margin-top:8px"><span>Total Deductions</span><span style="color:#d94f4f;font-variant-numeric:tabular-nums">AED {fmt(d["total"])}</span></div>
<div style="display:flex;justify-content:space-between;padding:14px 0 0;margin-top:12px;border-top:2px solid #1565a0"><span style="font-size:9px;letter-spacing:2px;text-transform:uppercase;color:#6b7280;align-self:center">Net Owner Payout</span><span style="font-size:22px;font-weight:800;color:#1565a0;font-variant-numeric:tabular-nums">AED {fmt(k["net_payout_exact"])}</span></div></div>
</div>
<div style="margin:0 44px 18px;padding:12px 20px;background:rgba(21,101,160,.07);border-left:3px solid #1565a0;border-radius:0 8px 8px 0;font-size:12px;color:#0d4a75"><strong>Payment Schedule:</strong> {month_short(m)} payout will be processed on <strong>28th {next_month_label(m)}</strong></div>
<div style="padding:12px 44px 20px">
<div style="font-size:9px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#1565a0;padding:0 0 12px;display:flex;align-items:center;gap:14px">Notes & Definitions<span style="flex:1;height:1px;background:#dde3ea"></span></div>
<div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px;line-height:1.6;color:#6b7280"><span style="color:#1565a0;font-weight:700;min-width:16px">1.</span><div><strong style="color:#1a1d24;font-weight:500">Booking Revenue:</strong> Total amount collected from the guest including accommodation, cleaning, tourism, VAT, and other fees.</div></div>
<div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px;line-height:1.6;color:#6b7280"><span style="color:#1565a0;font-weight:700;min-width:16px">2.</span><div><strong style="color:#1a1d24;font-weight:500">Commission:</strong> Platform host fees (Airbnb, Booking.com) and payment processing charges.</div></div>
<div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px;line-height:1.6;color:#6b7280"><span style="color:#1565a0;font-weight:700;min-width:16px">3.</span><div><strong style="color:#1a1d24;font-weight:500">Net Revenue:</strong> Amount remitted to Radiant Homes after platform and payment deductions.</div></div>
<div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px;line-height:1.6;color:#6b7280"><span style="color:#1565a0;font-weight:700;min-width:16px">4.</span><div><strong style="color:#1a1d24;font-weight:500">PM 15%:</strong> Property Management Commission calculated at 15% of revenue net of retained fees.</div></div>
<div style="display:flex;gap:10px;margin-bottom:7px;font-size:11px;line-height:1.6;color:#6b7280"><span style="color:#1565a0;font-weight:700;min-width:16px">5.</span><div><strong style="color:#1a1d24;font-weight:500">Owner Gross:</strong> Amount before operational expenses. Owner Gross less Expenses equals Net Owner Payout.</div></div>
</div>
<div style="padding:18px 44px;border-top:1px solid #dde3ea;display:flex;justify-content:space-between;font-size:10px;color:#6b7280"><span>Radiant Vacation Homes Rental L.L.C</span><span>3503, Aspect Tower, Business Bay, UAE</span></div>
</div></body></html>'''


# ─── PDF Generation ───────────────────────────────────────────────────────────

def html_to_pdf(html_content):
    """Convert HTML string to PDF bytes using Playwright."""
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8") as f:
        f.write(html_content)
        tmp_html = f.name

    tmp_pdf = tmp_html.replace(".html", ".pdf")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(f"file://{os.path.abspath(tmp_html)}", wait_until="networkidle")
            try:
                page.wait_for_function(
                    "() => document.fonts.ready.then(() => document.fonts.size > 0)",
                    timeout=10000,
                )
            except Exception:
                pass
            page.wait_for_timeout(1500)

            height = page.evaluate("() => document.querySelector('.page').scrollHeight")
            width = page.evaluate("() => document.querySelector('.page').offsetWidth")

            page.pdf(
                path=tmp_pdf,
                width=f"{width}px",
                height=f"{height + 20}px",
                print_background=True,
                margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
            )
            browser.close()

        with open(tmp_pdf, "rb") as f:
            return f.read()
    finally:
        for p in [tmp_html, tmp_pdf]:
            if os.path.exists(p):
                os.unlink(p)


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    """Parse workbook and return available units + months."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an .xlsx file"}), 400

    try:
        wb = openpyxl.load_workbook(io.BytesIO(file.read()), data_only=True)
        units = load_unit_registry(wb)
        months = get_available_months(wb, units)

        # Store workbook bytes in session-like temp file
        file.seek(0)
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        file.save(tmp.name)
        wb_path = tmp.name

        return jsonify({
            "wb_path": wb_path,
            "units": units,
            "months": months,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def generate():
    """Generate PDFs and return as a zip."""
    data = request.json
    wb_path = data.get("wb_path")
    month = data.get("month")
    unit_codes = data.get("units", [])
    logo_b64 = data.get("logo_b64")

    if not wb_path or not os.path.exists(wb_path):
        return jsonify({"error": "Workbook not found. Please re-upload."}), 400

    wb = openpyxl.load_workbook(wb_path, data_only=True)
    units = load_unit_registry(wb)

    zip_buffer = io.BytesIO()
    results = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for code in unit_codes:
            unit = next((u for u in units if u["code"] == code), None)
            if not unit:
                results.append({"code": code, "status": "error", "msg": "Unit not found"})
                continue

            pnl = load_pnl(wb, code, month)
            if not pnl:
                results.append({"code": code, "status": "skip", "msg": "No P&L data"})
                continue

            bookings = load_bookings(wb, code, month)
            if not bookings:
                results.append({"code": code, "status": "skip", "msg": "No bookings"})
                continue

            soa = calculate_soa(unit, pnl, bookings, month)
            html = generate_html(soa, logo_b64)
            pdf_bytes = html_to_pdf(html)

            unit_num = code.split(" ")[1] if " " in code else code
            filename = f"{unit['building'].replace(' ', '_')}_{unit_num}_SOA_{month_label(month).replace(' ', '_')}.pdf"
            zf.writestr(filename, pdf_bytes)

            results.append({
                "code": code,
                "status": "ok",
                "name": soa["property_name"],
                "payout": soa["kpi"]["net_payout_exact"],
                "gross": soa["kpi"]["owner_gross"],
                "bookings": soa["kpi"]["reservations"],
                "filename": filename,
            })

    zip_buffer.seek(0)

    # Save zip to temp file
    tmp_zip = tempfile.NamedTemporaryFile(suffix=".zip", delete=False)
    tmp_zip.write(zip_buffer.getvalue())
    tmp_zip.close()

    return jsonify({"results": results, "zip_path": tmp_zip.name})


@app.route("/api/download")
def download():
    """Download the generated zip file."""
    zip_path = request.args.get("path")
    if not zip_path or not os.path.exists(zip_path):
        return "File not found", 404

    month = request.args.get("month", "statements")
    return send_file(
        zip_path,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"Radiant_Homes_SOAs_{month.replace(' ', '_')}.zip",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
