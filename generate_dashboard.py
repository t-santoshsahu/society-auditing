#!/usr/bin/env python3
"""Generate a single static dashboard page (dashboard.html) from the monthly report_*.xlsx files.

Usage:
  python generate_dashboard.py [data_root] [-o output.html]

data_root is either one financial-year folder (e.g. test_data/2026-2027) or a folder holding several
(e.g. test_data). Every YYYY-YYYY folder containing */report_*.xlsx becomes an entry in the FY dropdown,
newest year first, so new years are picked up automatically.
"""
import argparse
import glob
import os
import re
import sys

import monthly_dashboard

FY_DIR = re.compile(r"^(\d{4})-(\d{4})$")


def has_reports(path):
    return bool(glob.glob(os.path.join(path, "*", "report_*.xlsx")))


def find_financial_years(root):
    """Return [(label, folder)] newest first, e.g. ("FY27", ".../2026-2027")."""
    root = os.path.abspath(root)
    if FY_DIR.match(os.path.basename(root)) or has_reports(root):
        candidates = [root]
    else:
        candidates = [
            os.path.join(root, name) for name in os.listdir(root)
            if FY_DIR.match(name) and os.path.isdir(os.path.join(root, name))
        ]
    found = []
    for folder in candidates:
        match = FY_DIR.match(os.path.basename(folder))
        if match and has_reports(folder):
            found.append((int(match.group(2)), "FY" + match.group(2)[2:], folder))
    return [(label, folder) for _, label, folder in sorted(found, reverse=True)]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("root", nargs="?", default="test_data", help="a financial-year folder or a folder containing several")
    parser.add_argument("-o", "--output", help="output HTML path (default: <root>/dashboard.html)")
    args = parser.parse_args()

    years = find_financial_years(args.root)
    if not years:
        sys.exit("No <YYYY-YYYY>/<month>/report_*.xlsx files found under %s" % args.root)

    datasets = {label: monthly_dashboard.load_year(folder) for label, folder in years}
    output = args.output or os.path.join(os.path.abspath(args.root), "dashboard.html")
    monthly_dashboard.write_dashboard(datasets, output)


if __name__ == "__main__":
    main()
