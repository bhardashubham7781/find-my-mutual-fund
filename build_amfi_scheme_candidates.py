#!/usr/bin/env python3
"""Build a review-only scheme master candidate list from AMFI's official NAV file.

This script intentionally DOES NOT activate schemes for publication. Every output
row has include_v1=false until scheme names/categories are reviewed and portfolio
matches are checked.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen

AMFI_NAV_URL = "https://portal.amfiindia.com/spages/NAVAll.txt"
TARGET_AMCS = {
    "sbi",
    "hdfc",
    "icici prudential",
    "nippon india",
    "kotak mahindra",
    "axis",
    "uti",
    "aditya birla sun life",
    "mirae asset",
    "dsp",
}
MAX_NAV_AGE_DAYS = 45
HEADERS = [
    "amc", "scheme_name", "amfi_scheme_code", "category", "plan", "option",
    "include_v1", "candidate_status", "amfi_section", "amfi_nav_date",
]


def clean(value: object) -> str:
    return "" if value is None else str(value).strip().lstrip("\ufeff")


def normalise(value: object) -> str:
    text = clean(value).lower().replace("&", " and ")
    text = text.replace("–", "-").replace("—", "-").replace("’", "'")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalise_amc(value: object) -> str:
    text = normalise(value)
    text = re.sub(r"\b(mutual fund|asset management company|asset management)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_date(value: str) -> date | None:
    value = clean(value)
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def extract_section(line: str) -> str | None:
    match = re.match(
        r"^(?:Open Ended Schemes|Close Ended Schemes|Interval Fund Schemes)\s*\((.*)\)\s*$",
        line.strip(), flags=re.IGNORECASE,
    )
    return match.group(1).strip() if match else None


def user_category(section: str, scheme_name: str) -> str | None:
    """Map AMFI category headers to this product's requested v1 category labels."""
    s = normalise(section)
    n = normalise(scheme_name)

    # Explicitly exclude combined categories that would otherwise substring-match.
    if "large and mid cap" in s or "large mid cap" in s:
        return None

    if "large cap fund" in s:
        return "Large Cap"
    if "mid cap fund" in s:
        return "Mid Cap"
    if "small cap fund" in s:
        return "Small Cap"
    if "flexi cap fund" in s:
        return "Flexi Cap"
    if "multi cap fund" in s:
        return "Multi Cap"
    if "elss" in s or "equity linked savings scheme" in s:
        return "ELSS"
    if "sectoral thematic fund" in s or "sectoral thematic" in s:
        if re.search(r"\b(psu|public sector)\b", n):
            return "PSU"
        return "Sectoral/Thematic"

    # AMFI places index schemes under a separate "Index Funds" grouping; do not
    # assume every such index is equity because debt index funds can appear there.
    if "index fund" in s:
        equity_index_markers = (
            "nifty", "sensex", "bse", "nse", "s and p 500", "sp 500",
            "msci", "nasdaq", "dow jones", "ftse", "hang seng",
        )
        debt_markers = (
            "g sec", "gsec", "government bond", "bond index", "debt index",
            "target maturity", "constant maturity", "money market", "liquid",
            "overnight", "corporate bond", "gilt", "treasury",
        )
        if any(marker in n for marker in debt_markers):
            return None
        if re.search(r"\betf\b", n):
            return None  # v1 is Direct-Growth mutual fund schemes, not ETFs
        if any(marker in n for marker in equity_index_markers):
            return "Index Fund"

    # Other equity categories are deliberately outside the stated initial scope.
    return None


def is_direct_growth(plan: str, option: str, scheme_name: str) -> bool:
    p = normalise(plan)
    o = normalise(option)
    n = normalise(scheme_name)
    direct = "direct plan" in p or p == "direct" or "direct growth" in o
    growth = "growth" in o or "direct growth" in o
    distribution_option = any(
        word in o for word in ("idcw", "dividend", "payout", "reinvestment", "bonus")
    )
    # Legacy 6-column AMFI files can put plan/option in the Scheme Name field.
    if not direct:
        direct = "direct plan" in n or re.search(r"\bdirect growth\b", n) is not None
    if not growth:
        growth = re.search(r"\bgrowth(?: option)?\b", n) is not None
    if not distribution_option:
        distribution_option = any(
            word in n for word in ("idcw", "dividend payout", "dividend reinvestment", "bonus option")
        )
    return direct and growth and not distribution_option


def clean_base_scheme_name(raw_name: str) -> str:
    base = clean(raw_name)
    # Strip plan/option suffixes from older NAV-file formats before normalising it.
    suffix_patterns = [
        r"\s*[-–]\s*(?:direct plan|regular plan|institutional plan|direct growth)\s*$",
        r"\s*[-–]\s*(?:growth option|growth)\s*$",
        r"\s*[-–]\s*(?:idcw(?:\s+.*)?|dividend(?:\s+.*)?|bonus(?:\s+option)?)\s*$",
    ]
    changed = True
    while changed:
        old = base
        for pattern in suffix_patterns:
            base = re.sub(pattern, "", base, flags=re.IGNORECASE).strip()
        changed = base != old
    return base


def complete_scheme_name(raw_name: str) -> str:
    base = clean_base_scheme_name(raw_name)
    return f"{base} - Direct Plan - Growth" if base else clean(raw_name)


def download_nav_file(url: str) -> str:
    request = Request(url, headers={"User-Agent": "FindMyMutualFund/1.0 (scheme metadata lookup)"})
    with urlopen(request, timeout=60) as response:
        payload = response.read()
    # AMFI currently publishes this feed as UTF-8 text; tolerate a BOM.
    return payload.decode("utf-8-sig", errors="replace")


def parse_nav_text(text: str, today: date | None = None) -> tuple[list[dict[str, str]], dict[str, object]]:
    today = today or date.today()
    current_amc: str | None = None
    current_section: str | None = None
    candidate_map: dict[str, dict[str, str]] = {}
    row_count = 0
    stale_count = 0
    malformed_count = 0
    non_target_amc_count = 0
    out_of_scope_count = 0
    non_direct_growth_count = 0
    no_date_count = 0

    for raw_line in text.splitlines():
        line = raw_line.strip().lstrip("\ufeff")
        if not line:
            continue
        section = extract_section(line)
        if section is not None:
            current_section = section
            current_amc = None
            continue
        if ";" not in line:
            if "scheme code" not in line.lower() and not line.lower().startswith("note:"):
                current_amc = line
            continue

        fields = [clean(part) for part in line.split(";")]
        if fields and fields[0].lower() in {"scheme code", "schemecode"}:
            continue
        if len(fields) not in (6, 8):
            malformed_count += 1
            continue
        row_count += 1

        if len(fields) == 8:
            scheme_code, _isin_growth, _isin_reinvestment, raw_name, plan, option, _nav, nav_date_text = fields
        else:
            scheme_code, _isin_growth, _isin_reinvestment, raw_name, _nav, nav_date_text = fields
            # Older feed format stores plan and option as a suffix in Scheme Name.
            lower_name = raw_name.lower()
            plan = "Direct Plan" if "direct plan" in lower_name or "direct growth" in lower_name else ""
            option = "Growth" if re.search(r"\bgrowth(?: option)?\b", lower_name) else ""

        amc_key = normalise_amc(current_amc or "")
        if amc_key not in TARGET_AMCS:
            non_target_amc_count += 1
            continue
        if not current_section:
            out_of_scope_count += 1
            continue

        category = user_category(current_section, raw_name)
        if category is None:
            out_of_scope_count += 1
            continue
        if not is_direct_growth(plan, option, raw_name):
            non_direct_growth_count += 1
            continue
        nav_date = parse_date(nav_date_text)
        if nav_date is None:
            no_date_count += 1
            continue
        age_days = (today - nav_date).days
        if age_days < -1 or age_days > MAX_NAV_AGE_DAYS:
            stale_count += 1
            continue
        if not scheme_code or not scheme_code.isdigit() or not raw_name:
            malformed_count += 1
            continue

        # Keep one candidate row per scheme code; scheme codes distinguish plan/options.
        candidate = {
            "amc": current_amc or "",
            "scheme_name": complete_scheme_name(raw_name),
            "amfi_scheme_code": scheme_code,
            "category": category,
            "plan": "Direct Plan",
            "option": "Growth",
            "include_v1": "false",
            "candidate_status": "REVIEW_REQUIRED_NOT_PUBLISHED",
            "amfi_section": current_section,
            "amfi_nav_date": nav_date.isoformat(),
        }
        candidate_map[scheme_code] = candidate

    candidates = sorted(
        candidate_map.values(),
        key=lambda row: (row["amc"].lower(), row["category"].lower(), row["scheme_name"].lower()),
    )
    summary: dict[str, object] = {
        "source_url": AMFI_NAV_URL,
        "generated_date": today.isoformat(),
        "max_nav_age_days": MAX_NAV_AGE_DAYS,
        "raw_nav_rows_examined": row_count,
        "candidate_rows": len(candidates),
        "candidates_by_amc": dict(sorted(Counter(row["amc"] for row in candidates).items())),
        "candidates_by_category": dict(sorted(Counter(row["category"] for row in candidates).items())),
        "excluded_non_target_amc_rows": non_target_amc_count,
        "excluded_out_of_scope_category_or_index_rows": out_of_scope_count,
        "excluded_non_direct_growth_rows_in_scope": non_direct_growth_count,
        "excluded_stale_nav_rows": stale_count,
        "excluded_rows_without_parseable_date": no_date_count,
        "malformed_row_count": malformed_count,
        "important": "This is a review-only candidate list. include_v1 is false for every row. Verify category scope and match each candidate to the AMC portfolio disclosure before setting include_v1=true.",
    }
    return candidates, summary


def main() -> int:
    parser = argparse.ArgumentParser(description="Create review-only Direct-Growth scheme candidates from AMFI NAVAll.txt.")
    parser.add_argument("--url", default=AMFI_NAV_URL, help="AMFI NAV feed URL (primarily for testing).")
    parser.add_argument("--output", type=Path, default=Path("reports/amfi_scheme_candidates.csv"))
    parser.add_argument("--summary", type=Path, default=Path("reports/amfi_scheme_candidates_summary.json"))
    args = parser.parse_args()

    try:
        text = download_nav_file(args.url)
    except (URLError, TimeoutError, OSError) as exc:
        print(f"ERROR: Could not download AMFI NAV feed from {args.url}: {exc}", file=sys.stderr)
        return 2

    candidates, summary = parse_nav_text(text)
    if not candidates:
        print("ERROR: No scheme candidates were found. Do not use an empty candidate list; inspect AMFI feed format.", file=sys.stderr)
        print(json.dumps(summary, indent=2), file=sys.stderr)
        return 1

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(candidates)
    args.summary.parent.mkdir(parents=True, exist_ok=True)
    args.summary.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print("AMFI candidate build completed (review-only; no funds activated).")
    print(f"Candidate CSV: {args.output}")
    print(f"Summary: {args.summary}")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
