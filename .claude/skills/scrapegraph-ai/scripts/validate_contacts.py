#!/usr/bin/env python3
"""Validate scraped company contact details before they reach an app.

An LLM scraper produces plausible-looking contact details; this turns "plausible" into
"checkable". It normalises values, checks they are coherent with the company's own domain,
counts independent corroboration, and assigns a publish/hold/reject status. Nothing here
calls a model or the network — it is deterministic so it can run in CI and so a failure is
always reproducible.

Input: JSON list of company records.

    [
      {
        "company": "Example Bank",
        "website": "https://www.examplebank.co.za",
        "alias_domains": ["examplebank.com"],
        "observations": [
          {"kind": "phone_complaints", "value": "0860 10 20 43",
           "source_url": "https://www.examplebank.co.za/complaints",
           "source_tier": 1, "observed_at": "2026-09-01T08:00:00Z"}
        ]
      }
    ]

Usage:
    python validate_contacts.py contacts.json                 # human-readable report
    python validate_contacts.py contacts.json --json out.json # machine-readable
    python validate_contacts.py --self-test                   # verify the checks work

Exit code is 1 if any observation is rejected, so it can gate a pipeline.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field, asdict
from datetime import datetime, timedelta, timezone
from typing import Any
from urllib.parse import urlparse

# Default country for bare national numbers. ZA = +27; change for other markets.
DEFAULT_COUNTRY_CODE = "27"
NATIONAL_DIGITS = 9  # digits after the country code for ZA numbers

# How long a verification stays fresh before the value must be re-confirmed.
FRESHNESS_DAYS = 45

PHONE_KINDS = {"phone_general", "phone_complaints", "phone_fraud", "phone_claims", "fax"}
EMAIL_KINDS = {"email_general", "email_complaints", "email_privacy"}
URL_KINDS = {"contact_url", "complaints_url", "website"}

EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")

# Multi-label suffixes we care about, so "co.za" is not mistaken for the registrable domain.
MULTI_LABEL_SUFFIXES = {
    "co.za", "org.za", "net.za", "web.za", "gov.za", "ac.za", "co.uk", "org.uk",
    "com.au", "co.nz", "com.br", "co.ke", "co.zw", "com.ng",
}

# Free-mail and generic hosts that are never a company's own contact domain.
GENERIC_EMAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.co.za", "outlook.com",
    "hotmail.com", "live.com", "icloud.com", "protonmail.com", "webmail.co.za",
    "mweb.co.za", "telkomsa.net", "aol.com",
}


def registrable_domain(host: str) -> str:
    """Return the registrable domain, handling multi-label suffixes like co.za.

    www.examplebank.co.za -> examplebank.co.za
    """
    host = (host or "").strip().lower().rstrip(".")
    if host.startswith("www."):
        host = host[4:]
    parts = host.split(".")
    if len(parts) < 2:
        return host
    last_two = ".".join(parts[-2:])
    if last_two in MULTI_LABEL_SUFFIXES and len(parts) >= 3:
        return ".".join(parts[-3:])
    return last_two


def domain_of_url(url: str) -> str:
    parsed = urlparse(url if "//" in (url or "") else f"//{url}")
    return registrable_domain(parsed.hostname or "")


def normalise_phone(raw: str) -> tuple[str | None, str | None]:
    """Normalise a phone number to E.164. Returns (normalised, error)."""
    if not raw or not raw.strip():
        return None, "empty"
    # Keep an intentional leading +, drop presentational characters.
    cleaned = raw.strip()
    has_plus = cleaned.startswith("+") or cleaned.startswith("00")
    digits = re.sub(r"\D", "", cleaned)
    if not digits:
        return None, "no digits"
    if cleaned.startswith("00"):
        digits = digits[2:]
    if digits.startswith(DEFAULT_COUNTRY_CODE) and len(digits) == len(DEFAULT_COUNTRY_CODE) + NATIONAL_DIGITS:
        national = digits[len(DEFAULT_COUNTRY_CODE):]
    elif digits.startswith("0") and len(digits) == NATIONAL_DIGITS + 1:
        national = digits[1:]
    elif len(digits) == NATIONAL_DIGITS and not has_plus:
        # Printed without the trunk 0, e.g. "860 10 20 43".
        national = digits
    else:
        if has_plus:
            return None, f"international number not in {DEFAULT_COUNTRY_CODE} space or wrong length ({len(digits)} digits)"
        return None, f"wrong length: {len(digits)} digits, expected {NATIONAL_DIGITS} national digits"
    if national.startswith("0"):
        return None, "national number starts with 0 after trunk stripping"
    return f"+{DEFAULT_COUNTRY_CODE}{national}", None


@dataclass
class Check:
    name: str
    passed: bool
    detail: str = ""


@dataclass
class Result:
    company: str
    kind: str
    value_as_printed: str
    normalised: str | None
    status: str                      # verified | needs_review | rejected
    checks: list[Check] = field(default_factory=list)
    corroborating_sources: list[str] = field(default_factory=list)
    reasons: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["checks"] = [asdict(c) for c in self.checks]
        return d


def _parse_ts(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def validate_company(record: dict[str, Any], now: datetime | None = None) -> list[Result]:
    now = now or datetime.now(timezone.utc)
    company = record.get("company", "<unnamed>")
    own_domains = {domain_of_url(record.get("website", ""))}
    own_domains |= {registrable_domain(d) for d in record.get("alias_domains", [])}
    own_domains.discard("")

    # Group observations by (kind, normalised value) so corroboration can be counted.
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    normalisation_errors: list[Result] = []

    for obs in record.get("observations", []):
        kind = obs.get("kind", "")
        raw = str(obs.get("value", ""))
        if kind in PHONE_KINDS:
            norm, err = normalise_phone(raw)
        else:
            norm, err = raw.strip(), None
        if err or not norm:
            normalisation_errors.append(
                Result(
                    company=company, kind=kind, value_as_printed=raw, normalised=None,
                    status="rejected",
                    checks=[Check("format", False, err or "empty value")],
                    reasons=[f"could not normalise: {err or 'empty value'}"],
                )
            )
            continue
        grouped.setdefault((kind, norm.lower()), []).append({**obs, "_normalised": norm})

    results: list[Result] = list(normalisation_errors)

    for (kind, _key), observations in grouped.items():
        normalised = observations[0]["_normalised"]
        printed = str(observations[0].get("value", ""))
        checks: list[Check] = [Check("format", True, "normalises cleanly")]
        reasons: list[str] = []

        # --- format, per kind -------------------------------------------------
        if kind in EMAIL_KINDS:
            ok = bool(EMAIL_RE.match(normalised))
            checks.append(Check("email_syntax", ok, "" if ok else "does not match email pattern"))
            if not ok:
                reasons.append("malformed email address")
            email_domain = registrable_domain(normalised.split("@")[-1]) if "@" in normalised else ""
            generic = email_domain in GENERIC_EMAIL_DOMAINS
            checks.append(Check("not_free_mail", not generic,
                                f"{email_domain} is a free-mail host" if generic else ""))
            if generic:
                reasons.append("free-mail address is not a company contact channel")
            coherent = bool(own_domains) and email_domain in own_domains
            checks.append(Check("domain_matches_company", coherent,
                                "" if coherent else f"{email_domain or '?'} not in {sorted(own_domains) or '[]'}"))
            if not coherent:
                reasons.append("email domain does not belong to the company")

        if kind in URL_KINDS:
            https = normalised.lower().startswith("https://")
            checks.append(Check("https", https, "" if https else "not an https URL"))
            if not https:
                reasons.append("contact URL is not https")
            same_site = domain_of_url(normalised) in own_domains if own_domains else False
            checks.append(Check("domain_matches_company", same_site,
                                "" if same_site else f"{domain_of_url(normalised)} not in {sorted(own_domains) or '[]'}"))
            if not same_site:
                reasons.append("contact URL is not on the company's own domain")

        # --- provenance -------------------------------------------------------
        tiers = [int(o.get("source_tier", 99)) for o in observations]
        best_tier = min(tiers) if tiers else 99
        tier1 = best_tier == 1
        checks.append(Check("tier1_source", tier1,
                            f"best source tier is {best_tier}" if not tier1 else ""))
        if not tier1:
            reasons.append("never seen on the company's own site or a statutory register")

        sources = sorted({str(o.get("source_url", "")) for o in observations if o.get("source_url")})
        missing_provenance = len(sources) < len(observations)
        checks.append(Check("has_source_urls", not missing_provenance,
                            "some observations have no source_url" if missing_provenance else ""))
        if missing_provenance:
            reasons.append("an observation is missing its source URL")

        independent = len({domain_of_url(s) for s in sources})
        corroborated = independent >= 2
        checks.append(Check("corroborated", corroborated,
                            f"{independent} independent source domain(s)"))
        if not corroborated:
            reasons.append("only one independent source — needs a second confirmation")

        # --- freshness --------------------------------------------------------
        stamps = [t for t in (_parse_ts(o.get("observed_at")) for o in observations) if t]
        newest = max(stamps) if stamps else None
        if newest is None:
            checks.append(Check("fresh", False, "no observed_at timestamp"))
            reasons.append("no timestamp, so freshness cannot be established")
        else:
            fresh = (now - newest) <= timedelta(days=FRESHNESS_DAYS)
            age = (now - newest).days
            checks.append(Check("fresh", fresh, f"last seen {age} day(s) ago"))
            if not fresh:
                reasons.append(f"last confirmed {age} days ago, past the {FRESHNESS_DAYS}-day window")

        # --- verdict ----------------------------------------------------------
        hard_failures = {"email_syntax", "domain_matches_company", "https", "not_free_mail",
                         "format", "has_source_urls"}
        failed = {c.name for c in checks if not c.passed}
        if failed & hard_failures or not tier1:
            status = "rejected" if (failed & hard_failures) else "needs_review"
        elif failed:
            status = "needs_review"
        else:
            status = "verified"

        results.append(Result(
            company=company, kind=kind, value_as_printed=printed, normalised=normalised,
            status=status, checks=checks, corroborating_sources=sources, reasons=reasons,
        ))

    return results


def validate_all(records: list[dict[str, Any]], now: datetime | None = None) -> list[Result]:
    out: list[Result] = []
    for record in records:
        out.extend(validate_company(record, now=now))
    return out


def print_report(results: list[Result]) -> None:
    symbols = {"verified": "PASS", "needs_review": "HOLD", "rejected": "FAIL"}
    by_company: dict[str, list[Result]] = {}
    for r in results:
        by_company.setdefault(r.company, []).append(r)

    for company, rows in by_company.items():
        print(f"\n{company}")
        print("-" * len(company))
        for r in rows:
            shown = r.normalised or r.value_as_printed
            print(f"  [{symbols[r.status]}] {r.kind:<20} {shown}")
            for reason in r.reasons:
                print(f"           - {reason}")

    counts = {k: sum(1 for r in results if r.status == k)
              for k in ("verified", "needs_review", "rejected")}
    print(f"\n{counts['verified']} publishable, {counts['needs_review']} held for review, "
          f"{counts['rejected']} rejected.")
    if counts["verified"] != len(results):
        print("Only 'publishable' rows should be shown in the app; link to the company's "
              "contact page for the rest.")


def self_test() -> int:
    """Confirm each check fires. Run this after editing the rules."""
    now = datetime(2026, 9, 1, tzinfo=timezone.utc)
    fresh = "2026-08-30T00:00:00Z"
    stale = "2025-01-01T00:00:00Z"

    record = {
        "company": "Example Bank",
        "website": "https://www.examplebank.co.za",
        "observations": [
            # Publishable: tier 1, two independent domains, fresh.
            {"kind": "phone_complaints", "value": "0860 10 20 43", "source_tier": 1,
             "source_url": "https://www.examplebank.co.za/complaints", "observed_at": fresh},
            {"kind": "phone_complaints", "value": "+27 86 010 2043", "source_tier": 2,
             "source_url": "https://www.obssa.co.za/members/examplebank", "observed_at": fresh},
            # Single source only -> hold.
            {"kind": "phone_general", "value": "021 941 1377", "source_tier": 1,
             "source_url": "https://www.examplebank.co.za/contact", "observed_at": fresh},
            # Free-mail on a bank page -> reject.
            {"kind": "email_complaints", "value": "examplebank.help@gmail.com", "source_tier": 1,
             "source_url": "https://www.examplebank.co.za/contact", "observed_at": fresh},
            # Directory-only -> hold (never publishable).
            {"kind": "phone_fraud", "value": "0800 123 456", "source_tier": 4,
             "source_url": "https://random-directory.com/examplebank", "observed_at": fresh},
            # Stale -> hold.
            {"kind": "contact_url", "value": "https://www.examplebank.co.za/contact",
             "source_tier": 1, "source_url": "https://www.examplebank.co.za/", "observed_at": stale},
            {"kind": "contact_url", "value": "https://www.examplebank.co.za/contact",
             "source_tier": 1, "source_url": "https://register.fsca.co.za/fsp/12345", "observed_at": stale},
            # Truncated number -> reject at normalisation.
            {"kind": "phone_claims", "value": "0860 10", "source_tier": 1,
             "source_url": "https://www.examplebank.co.za/claims", "observed_at": fresh},
        ],
    }

    results = {(r.kind, r.status) for r in validate_all([record], now=now)}
    expected = {
        ("phone_complaints", "verified"),
        ("phone_general", "needs_review"),
        ("email_complaints", "rejected"),
        ("phone_fraud", "needs_review"),
        ("contact_url", "needs_review"),
        ("phone_claims", "rejected"),
    }
    # Phone normalisation, checked independently of the pipeline.
    phone_cases = [
        ("0860 10 20 43", "+27860102043"),
        ("+27 86 010 2043", "+27860102043"),
        ("021 941 1377", "+27219411377"),
        ("0027219411377", "+27219411377"),
        ("(021) 941-1377", "+27219411377"),
    ]

    failures = []
    for missing in expected - results:
        failures.append(f"expected {missing} but did not get it; got {sorted(results)}")
    for raw, want in phone_cases:
        got, err = normalise_phone(raw)
        if got != want:
            failures.append(f"normalise_phone({raw!r}) -> {got!r} ({err}), expected {want!r}")
    for raw in ("0860 10", "12", ""):
        got, err = normalise_phone(raw)
        if got is not None:
            failures.append(f"normalise_phone({raw!r}) should have failed, got {got!r}")
    for host, want in (("www.examplebank.co.za", "examplebank.co.za"),
                       ("mail.sub.examplebank.co.za", "examplebank.co.za"),
                       ("examplebank.com", "examplebank.com")):
        if registrable_domain(host) != want:
            failures.append(f"registrable_domain({host!r}) -> {registrable_domain(host)!r}, expected {want!r}")

    if failures:
        print("SELF-TEST FAILED")
        for f in failures:
            print("  -", f)
        return 1
    print("Self-test passed: every check fires as intended.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", nargs="?", help="JSON file of company records")
    parser.add_argument("--json", dest="json_out", help="write results as JSON to this path")
    parser.add_argument("--self-test", action="store_true", help="verify the checks work")
    args = parser.parse_args()

    if args.self_test:
        return self_test()
    if not args.input:
        parser.error("an input file is required (or use --self-test)")

    with open(args.input) as fh:
        records = json.load(fh)
    if isinstance(records, dict):
        records = [records]

    results = validate_all(records)
    if args.json_out:
        with open(args.json_out, "w") as fh:
            json.dump([r.to_dict() for r in results], fh, indent=2)
        print(f"Wrote {len(results)} results to {args.json_out}")
    else:
        print_report(results)

    return 1 if any(r.status == "rejected" for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
