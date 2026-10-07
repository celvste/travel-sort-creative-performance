#!/usr/bin/env python3
"""
build_dashboard.py - bundle the creative-performance data into dashboard.html.

Sources, per campaign:
  workbooks/  master workbooks (..._D28_IAP_..._by_week.xlsx). Every weekly tab
              is read as-is; Master_Data is ignored.
  weeks/      tabs written by applovin_week.py. Added for any week the master
              workbook does not have yet.
  totals/     Total sheet CSV exports. Used as the Total tab only when the
              campaign has no master workbook here.

With a master workbook, the Total tab is the sum of every weekly tab shown, in
the Total sheet's row order. Re-run after each weekly pull, then republish.

    python3 build_dashboard.py
"""

import csv
import glob
import json
import os
import re
from datetime import date

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
TAB_RE = re.compile(r"^\d{2}\.\d{2}\.\d{2}[-~]\d{2}\.\d{2}\.\d{2}$")
TYPES = ["Video", "Playable", "Image"]
BLUE = "FFD6E4F5"            # applovin_week.py: Added Date from first serve
LABEL_ALIASES = {"BLD": "BLD iOS"}   # before Sep 23, BLD was the iOS campaign only
ORDER = ["IAP", "BLD iOS", "BLD Android", "ADS"]
DISPLAY = {"IAP": "IAP iOS", "ADS": "ADS iOS"}   # names shown on the page only


def num(value):
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value or "").replace("$", "").replace(",", "").replace("%", "").strip())
    except ValueError:
        return 0.0


def iso(value):
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    value = str(value or "").strip()
    return value if re.match(r"^\d{4}-\d{2}-\d{2}$", value) else None


def week_key(tab):
    return tab.replace("~", "-")


def read_tab(ws):
    rows = []
    for row in ws.iter_rows(min_row=2, max_col=8):
        name = row[2].value
        if not name:
            continue
        kind = str(row[0].value or "").strip() or "Video"
        flag = 0
        if row[1].value is None:
            flag = 2
        elif row[1].fill is not None and row[1].fill.fill_type and row[1].fill.fgColor.rgb == BLUE:
            flag = 1
        rows.append({"type": kind, "added": iso(row[1].value), "name": str(name).strip(),
                     "spend": num(row[3].value), "imp": int(num(row[5].value)),
                     "clk": int(num(row[6].value)), "flag": flag})
    return rows


def campaign_of(filename):
    m = re.search(r"\b(D\d+) (.+?) Applovin", filename)
    if not m:
        return None, None
    return m.group(1), LABEL_ALIASES.get(m.group(2), m.group(2))


def sum_weeks(order, weeks):
    agg = {}
    for tab in weeks.values():
        for r in tab:
            key = (r["type"], r["name"])
            a = agg.setdefault(key, {**r, "spend": 0.0, "imp": 0, "clk": 0, "flag": r["flag"]})
            a["spend"] += r["spend"]
            a["imp"] += r["imp"]
            a["clk"] += r["clk"]
            if a["added"] is None and r["added"]:
                a["added"], a["flag"] = r["added"], r["flag"]
    out, seen = [], set()
    for kind, name, added in order:
        a = agg.get((kind, name))
        if a and (kind, name) not in seen:
            out.append({**a, "added": added or a["added"], "flag": 0 if added else a["flag"]})
            seen.add((kind, name))
    for kind in TYPES:
        extra = sorted((a for k, a in agg.items() if k not in seen and a["type"] == kind),
                       key=lambda a: -a["spend"])
        last = max((i for i, r in enumerate(out) if r["type"] == kind), default=len(out) - 1)
        out[last + 1:last + 1] = extra
    return out


def main():
    camps = {}

    for path in sorted(glob.glob(os.path.join(HERE, "workbooks", "*.xlsx"))):
        window, label = campaign_of(os.path.basename(path))
        if not label:
            continue
        wb = openpyxl.load_workbook(path, data_only=True)
        weeks = {week_key(s): read_tab(wb[s]) for s in wb.sheetnames if TAB_RE.match(s)}
        order = []
        if "Total" in wb.sheetnames:
            for row in wb["Total"].iter_rows(min_row=2, max_col=3, values_only=True):
                if row[2]:
                    order.append((str(row[0]).strip(), str(row[2]).strip(), iso(row[1])))
        gap = None
        if "Master_Data" in wb.sheetnames:   # the workbook's Total sums Master_Data
            md = sum(num(r[3]) for r in wb["Master_Data"].iter_rows(min_row=2, values_only=True) if r[2])
            gap = round(sum(r["spend"] for t in weeks.values() for r in t) - md, 2)
        camps[label] = {"window": window, "source": "workbook", "weeks": weeks, "order": order, "gap": gap}

    for path in sorted(glob.glob(os.path.join(HERE, "weeks", "*", "*_Applovin_Creative_Performance_*.xlsx"))):
        week = os.path.basename(os.path.dirname(path))
        label = os.path.basename(path).split("_Applovin")[0]
        label = LABEL_ALIASES.get(label, label)
        c = camps.setdefault(label, {"window": None, "source": "pull", "weeks": {}, "order": []})
        if week not in c["weeks"]:
            c["weeks"][week] = read_tab(openpyxl.load_workbook(path).active)
            c.setdefault("pulled", []).append(week)

    for path in glob.glob(os.path.join(HERE, "totals", "*.csv")):
        window, label = campaign_of(os.path.basename(path))
        if not label or label not in camps:
            continue
        c = camps[label]
        c["window"] = c["window"] or window
        if c["source"] == "workbook":
            continue
        rows = []
        with open(path, newline="", encoding="utf-8-sig") as fh:
            reader = csv.reader(fh)
            next(reader)
            for r in reader:
                if len(r) >= 7 and r[2].strip():
                    rows.append({"type": r[0].strip(), "added": iso(r[1]), "name": r[2].strip(),
                                 "spend": num(r[3]), "imp": int(num(r[5])), "clk": int(num(r[6])),
                                 "flag": 0 if iso(r[1]) else 2})
        c["total_csv"] = rows

    # A blank Added Date in one week is often known from another week's pull of
    # the same campaign. Fill it from there (dated rows win over first-serve
    # ones, then the earliest date) and mark it blue, never touching the xlsx.
    for c in camps.values():
        tabs = list(c["weeks"].values()) + [c.get("total_csv", [])]
        known = {}
        for tab in tabs:
            for r in tab:
                if r["added"]:
                    best = known.get(r["name"])
                    if best is None or (r["flag"], r["added"]) < best:
                        known[r["name"]] = (r["flag"], r["added"])
        for tab in tabs:
            for r in tab:
                if not r["added"] and r["name"] in known:
                    r["added"], r["flag"] = known[r["name"]][1], 1

    # compact: names in one shared list, rows as arrays
    names, index = [], {}

    def pack(rows):
        out = []
        for r in rows:
            if r["name"] not in index:
                index[r["name"]] = len(names)
                names.append(r["name"])
            out.append([TYPES.index(r["type"]) if r["type"] in TYPES else 0, r["added"],
                        index[r["name"]], round(r["spend"], 2), r["imp"], r["clk"], r["flag"]])
        return out

    data = {"generated": date.today().isoformat(), "campaigns": [], "names": names}
    for label in sorted(camps, key=lambda l: (ORDER.index(l) if l in ORDER else 9, l)):
        c = camps[label]
        weeks = dict(sorted(c["weeks"].items(), reverse=True))
        if c["source"] == "workbook":
            total, total_src = sum_weeks(c["order"], weeks), "weeks"
        else:
            total, total_src = c.get("total_csv", []), "csv"
        data["campaigns"].append({
            "label": DISPLAY.get(label, label), "window": c["window"], "source": c["source"], "totalSource": total_src, "masterGap": c.get("gap"),
            "pulled": sorted(c.get("pulled", [])) if c["source"] == "workbook" else [],
            "sheets": [{"name": "Total", "rows": pack(total)}] +
                      [{"name": w, "rows": pack(rows)} for w, rows in weeks.items()],
        })

    template = open(os.path.join(HERE, "dashboard_template.html"), encoding="utf-8").read()
    payload = json.dumps(data, separators=(",", ":")).replace("</", "<\\/")
    out = os.path.join(HERE, "dashboard.html")
    open(out, "w", encoding="utf-8").write(template.replace("/*DATA*/null", payload))
    print(f"{out}  ({os.path.getsize(out) // 1024} KB)")
    for c in data["campaigns"]:
        print(f"  {c['label']:<12} {c['window'] or '':<4} {len(c['sheets']) - 1:>3} weeks  "
              f"Total from {c['totalSource']}  ${sum(r[3] for r in c['sheets'][0]['rows']):,.0f}")


if __name__ == "__main__":
    main()
