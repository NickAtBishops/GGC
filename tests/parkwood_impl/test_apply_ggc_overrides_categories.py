"""Deterministic tests for the category re-routing snap inside
`apply_ggc_overrides` (backend.py).

The pattern-based snap is the §0 / §5.4 safety net that locks in the
right GGC category before downstream rules (Omitt forcing, EGI build,
per-site overrides) consume the data. It runs at the top of
`apply_ggc_overrides` via `_snap_categories_by_pattern`, applying the
compiled regex rules in `_GGC_PATTERN_ROUTES` (backend.py around
line 4469). Because LLM judgment varies run-to-run on these exact lines
— "Janitorial" lands in Repair and Maintenance one run and Ground
Maintenance the next, etc. — the deterministic snap is what removes
that drift.

Each test builds a synthetic financials dict with a seller row that
intentionally carries the WRONG `ggcCategory`, runs
`apply_ggc_overrides`, and asserts the snap rewrote it to the canonical
target. Rent roll and property_info are kept minimal so the only
behavior under test is the category snap (no per-site overrides, no
mgmt-fee math, no taxes — those need a unit count and rent roll).
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
def _expense_row(seller_name: str, wrong_category: str,
                 t12: float = 1200.0) -> dict:
    """Synthetic expense row carrying a (deliberately wrong)
    `ggcCategory` so the test can verify the snap rewrote it."""
    return {
        "sellerName":      seller_name,
        "ggcCategory":     wrong_category,
        "fyPrior":         t12,
        "fyCurrent":       t12,
        "brokerProforma":  t12,
        "t12Total":        t12,
        "monthly":         [t12 / 12] * 12,
        "ggcUnderwritten": t12,
        "confidence":      "medium",
        "notes":           "",
    }


def _income_row(seller_name: str, wrong_category: str,
                t12: float = 2400.0) -> dict:
    return {
        "sellerName":      seller_name,
        "ggcCategory":     wrong_category,
        "fyPrior":         t12,
        "fyCurrent":       t12,
        "brokerProforma":  t12,
        "t12Total":        t12,
        "monthly":         [t12 / 12] * 12,
        "ggcUnderwritten": t12,
        "confidence":      "medium",
        "notes":           "",
    }


def _make_financials(*, income: list[dict] | None = None,
                     expenses: list[dict] | None = None) -> dict:
    """Build a minimal financials shell. Empty rent roll + no units in
    property_info means downstream per-site / mgmt-fee / capex / tax
    overrides are no-ops, so the only behavior actually exercised is the
    category snap at the top of `apply_ggc_overrides`."""
    return {
        "income":   income or [],
        "expenses": expenses or [],
        "rentRoll": {
            "totalUnits":    0,
            "occupiedUnits": 0,
            "vacantUnits":   0,
            "occupancyRate": 0.0,
            "unitGroups":    [],
            "rentRollRows":  [],
        },
    }


def _find(rows: list[dict], seller_name: str) -> dict:
    matches = [r for r in rows if r.get("sellerName") == seller_name]
    assert matches, (
        f"Row with sellerName={seller_name!r} not found after "
        f"apply_ggc_overrides. Available: "
        f"{[r.get('sellerName') for r in rows]}"
    )
    assert len(matches) == 1, (
        f"Multiple rows match sellerName={seller_name!r}: {matches}"
    )
    return matches[0]


# --------------------------------------------------------------------------- #
# Expense-side snaps                                                          #
# --------------------------------------------------------------------------- #
def test_janitorial_snaps_to_ground_maintenance():
    """`Janitorial` rows must land in Ground Maintenance, not the
    generic Repair and Maintenance bucket the LLM commonly picks."""
    financials = _make_financials(expenses=[
        _expense_row("Janitorial", "Repair and Maintenance"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["expenses"], "Janitorial")
    assert row["ggcCategory"] == "Ground Maintenance", row


def test_equipment_fuel_snaps_to_gas_fuel():
    """`Equipment Fuel` is a fuel line — Gas/Fuel, not Repair and
    Maintenance / vehicle Omitt."""
    financials = _make_financials(expenses=[
        _expense_row("Equipment Fuel", "Repair and Maintenance"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["expenses"], "Equipment Fuel")
    assert row["ggcCategory"] == "Gas/Fuel", row


def test_gifts_to_residents_snaps_to_omitt_expense():
    """Resident gifts are one-time discretionary spend → Omitt Expense
    so they don't pollute OpEx / NOI."""
    financials = _make_financials(expenses=[
        _expense_row("Gifts to Residents", "G&A"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["expenses"], "Gifts to Residents")
    assert row["ggcCategory"] == "Omitt Expense", row


def test_tax_and_title_for_mobile_home_snaps_to_omitt_expense():
    """Home title-transfer costs are inventory-side, not opex. The
    seller often books them as RE Taxes — the snap rescues them to
    Omitt Expense."""
    financials = _make_financials(expenses=[
        _expense_row("Tax & Title for Mobile Home", "RE Taxes"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["expenses"], "Tax & Title for Mobile Home")
    assert row["ggcCategory"] == "Omitt Expense", row


def test_misc_snaps_to_ga():
    """A bare `Misc` line is GGC's G&A by convention, not Other."""
    financials = _make_financials(expenses=[
        _expense_row("Misc", "Other"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["expenses"], "Misc")
    assert row["ggcCategory"] == "G&A", row


def test_picnic_snaps_to_omitt_expense():
    """Resident picnic / one-time events → Omitt Expense (matches the
    later `_EXPENSE_NAME_PATTERNS` rule too — both snaps must agree)."""
    financials = _make_financials(expenses=[
        _expense_row("Picnic", "G&A"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["expenses"], "Picnic")
    assert row["ggcCategory"] == "Omitt Expense", row


# --------------------------------------------------------------------------- #
# Income-side snaps                                                           #
# --------------------------------------------------------------------------- #
def test_sales_tax_electric_income_snaps_to_utility_reimbursement():
    """`Sales Tax-Electric` is an electric-utility tax pass-through, so
    GGC classifies it as Utility Reimbursement, not Other Income."""
    financials = _make_financials(income=[
        _income_row("Sales Tax-Electric", "Other Income"),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["income"], "Sales Tax-Electric")
    assert row["ggcCategory"] == "Utility Reimbursement", row


def test_lc_payment_swanson_income_snaps_to_lease_to_own_income():
    """`LC Payment - Swanson` (or any `LC Payment - <tenant>` line) is
    a land-contract receivable. Single-row case: the consolidation pass
    needs ≥2 LC rows to roll up, so with one positive row the snap
    leaves it as a stand-alone `LTO` row."""
    financials = _make_financials(income=[
        _income_row("LC Payment - Swanson", "Other Income", t12=2400.0),
    ])

    backend.apply_ggc_overrides(financials, {})

    row = _find(financials["income"], "LC Payment - Swanson")
    assert row["ggcCategory"] == "LTO", row
    # And it must still be positive — sign should not have flipped.
    assert row["t12Total"] == pytest.approx(2400.0), row


# --------------------------------------------------------------------------- #
# Sanity: snap must record itself on `_ggcOverrides`                          #
# --------------------------------------------------------------------------- #
def test_snap_records_override_event():
    """Every category change must show up on `_ggcOverrides` so the
    Extraction Check tab can show the reviewer exactly what was
    rerouted. The before/after pair is what makes the run auditable."""
    financials = _make_financials(expenses=[
        _expense_row("Janitorial", "Repair and Maintenance"),
    ])

    backend.apply_ggc_overrides(financials, {})

    events = financials.get("_ggcOverrides") or []
    matched = [
        e for e in events
        if e.get("before") == "Repair and Maintenance"
        and e.get("after") == "Ground Maintenance"
        and "Janitorial" in (e.get("category") or "")
    ]
    assert matched, (
        f"Expected an override event recording Janitorial: "
        f"Repair and Maintenance → Ground Maintenance. Got: {events}"
    )
