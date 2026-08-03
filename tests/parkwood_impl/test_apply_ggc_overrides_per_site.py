"""Deterministic tests for the per-site flat overrides applied inside
`apply_ggc_overrides` (backend.py).

CONTEXT (CLAUDE.md §5.4 / OUTLINE root cause #6)
-------------------------------------------------
The per-site overrides are the deal-grade `$/unit/year × units` convention
that lands Underwriting!I35 at the canonical CorrectOutput values on every
run. These are GGC RULES (not LLM judgment) — they replace whatever the
methodology emitted, and they must be exact every time.

The defaults under test are taken from `GGC_PER_SITE_DEFAULTS` in
backend.py around line 4435:

  * Payroll              $425/unit/yr
  * Insurance (no flood) $250/unit/yr
  * Insurance (flood)    $300/unit/yr
  * Ground Maintenance   $200/unit/yr
  * G&A                  $100/unit/yr
  * Professional Fees    $ 50/unit/yr
  * Advertising          $  0/unit/yr   (default — inserted row suppressed)
  * Cap-Ex Reserve       $ 50/unit/yr   (CorrectOutput Parkwood value;
                                         the §5.4 capex pass runs first at
                                         $75 but the per-site pass runs
                                         AFTER and is the final word)

Bad Debt is synthesized at -2% × UW GPR per the bad-debt UW plug in
`_apply_per_site_overrides`.

WHAT THESE TESTS PIN
--------------------
For a 100-unit property with synthetic financials that already carry the
target categories (so the per-site overrides apply by category match,
not by reroute):

  Payroll              ggcUnderwritten == 425 × 100 == 42,500
  Insurance (no flood)                  == 250 × 100 == 25,000
  Insurance (flood)                     == 300 × 100 == 30,000
  Ground Maintenance                    == 200 × 100 == 20,000
  G&A                                   == 100 × 100 == 10,000
  Professional Fees                     ==  50 × 100 ==  5,000
  Advertising                           ==   0 × 100 ==      0
  Cap-Ex Reserve                        ==  50 × 100 ==  5,000
  Bad Debt (synthesized)                == -0.02 × UW GPR
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

# Make backend.py importable when pytest runs from repo root or elsewhere.
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(REPO_ROOT))

import backend  # noqa: E402


# --------------------------------------------------------------------------- #
# Fixture builders                                                            #
# --------------------------------------------------------------------------- #
def _expense_row(seller_name: str, ggc_category: str,
                 t12: float = 1200.0) -> dict:
    """Synthetic expense row. `t12Total` is intentionally non-zero so we
    can verify the per-site override REPLACED whatever number was there
    rather than appending on top of it. The methodology pass would have
    set `ggcUnderwritten` to the LLM's number — we mimic that here by
    seeding the field with the seller's T12 value."""
    return {
        "sellerName":      seller_name,
        "ggcCategory":     ggc_category,
        "fyPrior":         t12,
        "fyCurrent":       t12,
        "brokerProforma":  t12,
        "t12Total":        t12,
        "monthly":         [t12 / 12] * 12,
        "ggcUnderwritten": t12,
        "confidence":      "medium",
        "notes":           "",
    }


def _gpr_row(annual: float) -> dict:
    """Single Gross Potential Rent income row so the Bad Debt UW plug has
    a non-zero UW GPR to multiply -2% against."""
    return {
        "sellerName":      "Lot Rent Income",
        "ggcCategory":     "Gross Potential Rent",
        "fyPrior":         annual,
        "fyCurrent":       annual,
        "brokerProforma":  annual,
        "t12Total":        annual,
        "monthly":         [annual / 12] * 12,
        "ggcUnderwritten": annual,
        "confidence":      "high",
        "notes":           "",
    }


def _make_financials(*, gpr_annual: float = 500_000.0) -> dict:
    """Build a financials shell with every per-site category present on
    the expense side (so the override hits the by-category match path,
    not the inserted-row path) plus one GPR row so Bad Debt has something
    to multiply."""
    return {
        "income":   [_gpr_row(gpr_annual)],
        "expenses": [
            _expense_row("Payroll",            "Payroll"),
            _expense_row("Insurance Premium",  "Insurance"),
            _expense_row("Lawn / Grounds",     "Ground Maintenance"),
            _expense_row("Office & Admin",     "G&A"),
            _expense_row("Legal & Accounting", "Professional Fees"),
            _expense_row("Advertising / Mktg", "Advertising"),
            _expense_row("CapEx Reserve",      "Cap-Ex Reserve"),
        ],
        "rentRoll": {
            "totalUnits":    0,
            "occupiedUnits": 0,
            "vacantUnits":   0,
            "occupancyRate": 0.0,
            "unitGroups":    [],
            "rentRollRows":  [],
        },
    }


def _property_info(*, units: int = 100, flood: bool = False,
                   **extra) -> dict:
    """Form values as `apply_ggc_overrides` consumes them. `flood_zone`
    is read by `_apply_per_site_overrides` via `property_info["floodZone"]`
    (camelCase key) — backend.py around line 4626 — so we set the key
    the override layer actually reads, not the form field name."""
    pi: dict = {"units": units}
    if flood:
        pi["floodZone"] = "yes"
    pi.update(extra)
    return pi


def _find_by_category(rows: list[dict], category: str) -> dict:
    matches = [r for r in rows
               if (r.get("ggcCategory") or "").strip() == category]
    assert matches, (
        f"No row found with ggcCategory={category!r} after "
        f"apply_ggc_overrides. Categories present: "
        f"{[r.get('ggcCategory') for r in rows]}"
    )
    assert len(matches) == 1, (
        f"Multiple rows for ggcCategory={category!r}: {matches}"
    )
    return matches[0]


# --------------------------------------------------------------------------- #
# Per-site flat overrides on expense lines                                    #
# --------------------------------------------------------------------------- #
def test_payroll_per_site_default_425_times_100_units():
    """Payroll: $425/unit/yr × 100 units = $42,500. No form override on
    `payroll_per_site`, so the GGC default kicks in."""
    financials = _make_financials()
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Payroll")
    assert row["ggcUnderwritten"] == pytest.approx(42_500.0), row
    # Monthly history is now PRESERVED — the override only changes the
    # underwritten total. Previously the override clobbered seller
    # monthly history with a flat [target/12]*12 that destroyed
    # T3/T6/T12 trend bands. The fixture seeds monthly = [100]*12
    # (from `_expense_row`'s t12=1200), so we expect that to survive.
    assert row["monthly"] == pytest.approx([1_200.0 / 12] * 12), row


def test_insurance_per_site_non_flood_250_times_100_units():
    """Insurance, no flood: $250/unit/yr × 100 = $25,000."""
    financials = _make_financials()
    property_info = _property_info(units=100, flood=False)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Insurance")
    assert row["ggcUnderwritten"] == pytest.approx(25_000.0), row


def test_insurance_per_site_flood_zone_300_times_100_units():
    """Insurance, flood zone toggled on: $300/unit/yr × 100 = $30,000.
    Verifies the `floodZone` switch flips the default from $250 to $300
    inside `_apply_per_site_overrides` (the §5.4 insurance pass also
    multiplies by 1.15 in flood mode, but the per-site override runs
    AFTER and is the final word)."""
    financials = _make_financials()
    property_info = _property_info(units=100, flood=True)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Insurance")
    assert row["ggcUnderwritten"] == pytest.approx(30_000.0), row


def test_ground_maintenance_per_site_default_200_times_100_units():
    """Ground Maintenance: $200/unit/yr × 100 = $20,000."""
    financials = _make_financials()
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Ground Maintenance")
    assert row["ggcUnderwritten"] == pytest.approx(20_000.0), row


def test_ga_per_site_default_100_times_100_units():
    """G&A: $100/unit/yr × 100 = $10,000."""
    financials = _make_financials()
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "G&A")
    assert row["ggcUnderwritten"] == pytest.approx(10_000.0), row


def test_professional_fees_per_site_default_50_times_100_units():
    """Professional Fees: $50/unit/yr × 100 = $5,000."""
    financials = _make_financials()
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Professional Fees")
    assert row["ggcUnderwritten"] == pytest.approx(5_000.0), row


def test_advertising_per_site_default_zero_times_100_units():
    """Advertising: $0/unit/yr × 100 = $0. The seller's $1,200 t12 line
    must be REPLACED, not preserved — the per-site override is the rule.
    Note: when no Advertising row exists at all the override layer
    intentionally skips inserting a $0 row (backend.py around
    line 4703-4706: 'injecting a $0 row is just noise'). Here we DO
    have a seller-supplied Advertising row, so the override overwrites
    it to $0."""
    financials = _make_financials()
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Advertising")
    assert row["ggcUnderwritten"] == pytest.approx(0.0), row


def test_capex_per_site_default_50_times_100_units():
    """Cap-Ex Reserve: $50/unit/yr × 100 = $5,000. The §5.4 capex pass
    earlier in `apply_ggc_overrides` would have set this to $75 × 100 =
    $7,500, but `_apply_per_site_overrides` runs AFTER and reasserts the
    CorrectOutput-Parkwood $50/site value — that's the final word."""
    financials = _make_financials()
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    row = _find_by_category(financials["expenses"], "Cap-Ex Reserve")
    assert row["ggcUnderwritten"] == pytest.approx(5_000.0), row


# --------------------------------------------------------------------------- #
# Bad Debt UW plug                                                            #
# --------------------------------------------------------------------------- #
def test_bad_debt_synthesized_at_negative_two_percent_of_uw_gpr():
    """Bad Debt is synthesized by `_apply_per_site_overrides` as
    -bad_debt_uw_pct × UW GPR, defaulting to -2%. The fixture seeds a
    single GPR row at $500,000, so the synthesized Bad Debt must equal
    -0.02 × 500,000 = -$10,000 (and the row must be NEGATIVE, per the
    GGC sign convention pinned in §8 of CLAUDE.md)."""
    gpr_annual = 500_000.0
    financials = _make_financials(gpr_annual=gpr_annual)
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    bd_rows = [it for it in financials["income"]
               if (it.get("ggcCategory") or "").strip() == "Bad Debt"]
    assert bd_rows, (
        "Expected a synthesized Bad Debt row in income[] after "
        "apply_ggc_overrides. Categories present: "
        f"{[it.get('ggcCategory') for it in financials['income']]}"
    )
    assert len(bd_rows) == 1, bd_rows
    bd = bd_rows[0]

    # UW GPR after overrides — Bad Debt math uses the post-override GPR
    # because `_apply_per_site_overrides` runs last and reads
    # ggcUnderwritten on every GPR row. We seeded a single GPR row at
    # gpr_annual, so UW GPR == gpr_annual.
    uw_gpr = sum(
        float(it.get("ggcUnderwritten") or 0)
        for it in financials["income"]
        if (it.get("ggcCategory") or "").strip() == "Gross Potential Rent"
    )
    expected = -round(uw_gpr * 0.02, 2)
    assert uw_gpr == pytest.approx(gpr_annual), (
        f"GPR was unexpectedly mutated by apply_ggc_overrides: "
        f"got {uw_gpr}, expected {gpr_annual}"
    )
    assert bd["ggcUnderwritten"] == pytest.approx(expected), bd
    assert bd["ggcUnderwritten"] < 0, (
        f"Bad Debt must be NEGATIVE per GGC sign convention: got {bd}"
    )
    # And the per-site override layer should have flagged the synthesis.
    assert "Bad Debt" in (bd.get("sellerName") or "") or \
           "synthesized" in (bd.get("notes") or "").lower(), bd


def test_bad_debt_not_replugged_when_tier_c_already_zeroed_it():
    """Regression test: on Parkwood-shaped deals, realized credit losses
    (negative KCA Ventures / LC Payment rows) get netted into the
    consolidated GPR row upstream, and financials["_implicitCreditLoss"]
    records that amount. Tier C (inside apply_ggc_overrides, ahead of
    `_apply_per_site_overrides`) zeroes any existing Bad Debt line when
    that implicit loss is >= $10K, to avoid double-counting.

    Previously `_apply_per_site_overrides`'s bad-debt UW plug ran
    unconditionally right after and re-plugged Bad Debt to -2% x UW GPR,
    silently undoing Tier C's zeroing and double-counting the realized
    losses. This pins that the plug is skipped whenever Tier C's implicit
    -credit-loss signal is present, regardless of whether a Bad Debt row
    already existed."""
    gpr_annual = 500_000.0
    financials = _make_financials(gpr_annual=gpr_annual)
    # Seller-supplied Bad Debt row that Tier C should zero.
    financials["income"].append({
        "sellerName":      "Bad Debt / Write-offs",
        "ggcCategory":     "Bad Debt",
        "fyPrior":         -8_000.0, "fyCurrent": -8_000.0,
        "brokerProforma":  -8_000.0, "t12Total": -8_000.0,
        "monthly":         [-8_000.0 / 12] * 12,
        "ggcUnderwritten": -8_000.0,
        "confidence":      "medium",
        "notes":           "",
    })
    # Simulate the pre-consolidation capture of realized LC credit losses
    # (Parkwood: 40+ negative KCA Ventures rows) already netted into GPR.
    financials["_implicitCreditLoss"] = 25_000.0
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    bd_rows = [it for it in financials["income"]
               if (it.get("ggcCategory") or "").strip() == "Bad Debt"]
    assert len(bd_rows) == 1, bd_rows
    bd = bd_rows[0]
    assert bd["ggcUnderwritten"] == 0, (
        "Bad Debt should stay zeroed by Tier C — the per-site bad-debt "
        f"plug must not re-plug it when _implicitCreditLoss >= $10K. Got: {bd}"
    )
    assert bd["t12Total"] == 0, bd


def test_bad_debt_plug_still_fires_when_no_implicit_credit_loss():
    """Sanity check for the guard above: when `_implicitCreditLoss` is
    absent/zero (the normal, non-Parkwood-LC-contract case), the per-site
    bad-debt UW plug must still fire exactly as before."""
    gpr_annual = 500_000.0
    financials = _make_financials(gpr_annual=gpr_annual)
    property_info = _property_info(units=100)

    backend.apply_ggc_overrides(financials, property_info)

    bd_rows = [it for it in financials["income"]
               if (it.get("ggcCategory") or "").strip() == "Bad Debt"]
    assert len(bd_rows) == 1, bd_rows
    assert bd_rows[0]["ggcUnderwritten"] == pytest.approx(-10_000.0), bd_rows[0]
