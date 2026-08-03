"""Deterministic tests for `_ensure_rent_roll_complete` (backend.py).

The function is the §2.3 / §5.1 vacant-pad imputation backstop: when the
methodology returns fewer rent-roll rows than the user-stated unit count,
it appends synthetic "Vacant Lots" rows priced at per-type market lot rent
so GPR doesn't get silently undercounted.

The test case below exercises the canonical Parkwood-style shape: 97 rent
roll rows across a mix of TOH / POH / LTO / Flourish, with the property
form input saying units=100. After the call there must be 100 rows, the
three new rows must carry the expected sentinel values, and the
unitGroups vacant counts must also reflect the imputation.
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
# The dominant occupied unit type drives the imputed-row unitType and the
# imputed lotRent (mean of occupied-row lotRent values of that type). We
# build the fixture so TOH MH Site is unambiguously dominant: 60 rows at
# $500/mo. The other types are present but smaller and at different rents
# so a bug that picked the wrong dominant type would change the test
# verdict.
TOH_COUNT = 60
TOH_RENT = 500.0
POH_COUNT = 20
POH_RENT = 700.0
LTO_COUNT = 10
LTO_RENT = 800.0
FLOURISH_COUNT = 7
FLOURISH_RENT = 600.0

OCCUPIED_TOTAL = TOH_COUNT + POH_COUNT + LTO_COUNT + FLOURISH_COUNT  # 97
STATED_UNITS = 100
SHORTFALL = STATED_UNITS - OCCUPIED_TOTAL  # 3


def _make_row(seq: int, unit_type: str, lot_rent: float) -> dict:
    return {
        "tenantName": f"Tenant {seq}",
        "unitId":     str(seq),
        "unitType":   unit_type,
        "sellerType": unit_type,
        "status":     "Occupied",
        "lotRent":    lot_rent,
        "homeRent":   0,
        "lcPayment":  0,
        "moveInDate": "2020-01-01",
    }


def _build_financials() -> dict:
    rows: list[dict] = []
    seq = 1
    for _ in range(TOH_COUNT):
        rows.append(_make_row(seq, "TOH MH Site", TOH_RENT))
        seq += 1
    for _ in range(POH_COUNT):
        rows.append(_make_row(seq, "POH MH Site", POH_RENT))
        seq += 1
    for _ in range(LTO_COUNT):
        rows.append(_make_row(seq, "LTO MH Site", LTO_RENT))
        seq += 1
    for _ in range(FLOURISH_COUNT):
        rows.append(_make_row(seq, "Flourish MH Site", FLOURISH_RENT))
        seq += 1

    unit_groups = [
        {
            "unitType":      "TOH MH Site",
            "occupiedCount": TOH_COUNT,
            "vacantCount":   0,
            "lotRent":       TOH_RENT,
        },
        {
            "unitType":      "POH MH Site",
            "occupiedCount": POH_COUNT,
            "vacantCount":   0,
            "lotRent":       POH_RENT,
        },
        {
            "unitType":      "LTO MH Site",
            "occupiedCount": LTO_COUNT,
            "vacantCount":   0,
            "lotRent":       LTO_RENT,
        },
        {
            "unitType":      "Flourish MH Site",
            "occupiedCount": FLOURISH_COUNT,
            "vacantCount":   0,
            "lotRent":       FLOURISH_RENT,
        },
    ]

    return {
        "rentRoll": {
            "totalUnits":    OCCUPIED_TOTAL,
            "occupiedUnits": OCCUPIED_TOTAL,
            "vacantUnits":   0,
            "occupancyRate": 1.0,
            "unitGroups":    unit_groups,
            "rentRollRows":  rows,
        }
    }


# --------------------------------------------------------------------------- #
# Tests                                                                       #
# --------------------------------------------------------------------------- #
def test_appends_rows_to_match_stated_unit_count():
    """97 input rows + units=100 → exactly 3 rows appended."""
    financials = _build_financials()
    property_info = {"units": STATED_UNITS}

    backend._ensure_rent_roll_complete(financials, property_info)

    rows = financials["rentRoll"]["rentRollRows"]
    assert len(rows) == STATED_UNITS, (
        f"Expected {STATED_UNITS} rows after imputation, got {len(rows)}"
    )


def test_appended_rows_carry_vacant_sentinel_values():
    """The 3 new rows must be tenantName='Vacant Lots', status='Vacant',
    unitType = dominant occupied type (TOH MH Site here), and lotRent =
    average lotRent of TOH occupied rows ($500/mo)."""
    financials = _build_financials()
    property_info = {"units": STATED_UNITS}

    backend._ensure_rent_roll_complete(financials, property_info)

    rows = financials["rentRoll"]["rentRollRows"]
    new_rows = rows[-SHORTFALL:]

    assert len(new_rows) == SHORTFALL

    for row in new_rows:
        assert row["tenantName"] == "Vacant Lots", row
        assert row["status"] == "Vacant", row
        assert row["unitType"] == "TOH MH Site", row
        assert row["sellerType"] == "TOH MH Site", row
        # Dominant type is TOH (60 occupied rows), avg lotRent of those is
        # exactly TOH_RENT since every TOH row is at TOH_RENT.
        assert row["lotRent"] == pytest.approx(TOH_RENT), row
        assert row["homeRent"] == 0, row
        assert row["lcPayment"] == 0, row
        # Unit IDs must be unique and not collide with the 97 existing IDs.
        assert row["unitId"], row

    new_ids = {r["unitId"] for r in new_rows}
    existing_ids = {r["unitId"] for r in rows[:-SHORTFALL]}
    assert new_ids.isdisjoint(existing_ids), (
        f"Imputed unit IDs collide with existing IDs: "
        f"{new_ids & existing_ids}"
    )
    assert len(new_ids) == SHORTFALL, "Imputed unit IDs are not unique"


def test_unit_groups_counts_updated():
    """unitGroups must also see the shortfall distributed into vacantCount;
    new total across groups must equal stated_units."""
    financials = _build_financials()
    property_info = {"units": STATED_UNITS}

    backend._ensure_rent_roll_complete(financials, property_info)

    rr = financials["rentRoll"]
    groups = rr["unitGroups"]

    new_total = sum(
        (g.get("occupiedCount") or 0) + (g.get("vacantCount") or 0)
        for g in groups
    )
    assert new_total == STATED_UNITS, (
        f"unitGroups total {new_total} != stated_units {STATED_UNITS}"
    )

    total_vacant = sum(g.get("vacantCount") or 0 for g in groups)
    assert total_vacant == SHORTFALL, (
        f"unitGroups vacantCount sum {total_vacant} != shortfall {SHORTFALL}"
    )

    # rentRoll-level counters should mirror the recomputed totals.
    assert rr["totalUnits"] == STATED_UNITS
    assert rr["occupiedUnits"] == OCCUPIED_TOTAL
    assert rr["vacantUnits"] == SHORTFALL
    assert rr["occupancyRate"] == pytest.approx(
        OCCUPIED_TOTAL / STATED_UNITS
    )


def test_flag_and_extraction_check_emitted():
    """The function surfaces the imputation as both a high-severity flag
    and a WARN entry on the Extraction Check tab — that's the §0
    flag-and-prompt behavior the workbook reviewer relies on."""
    financials = _build_financials()
    property_info = {"units": STATED_UNITS}

    backend._ensure_rent_roll_complete(financials, property_info)

    flags = financials.get("flags") or []
    assert any(
        f.get("item") == "Rent roll vs unit count"
        and f.get("severity") == "high"
        for f in flags
    ), f"Expected a high-severity 'Rent roll vs unit count' flag; got {flags}"

    checks = financials.get("_extractionChecks") or []
    assert any(
        c.get("item") == "Vacant-pad imputation" and c.get("status") == "warn"
        for c in checks
    ), f"Expected a WARN 'Vacant-pad imputation' check; got {checks}"


def test_message_reports_rows_shortfall_not_groups_shortfall():
    """Regression test: when unitGroups already tally to stated_units
    (groups_shortfall=0) but rentRollRows is short (rows_shortfall>0), the
    reviewer-facing flag/check/print text must report the actual number of
    rows imputed, not 0. Previously these messages always interpolated
    `shortfall` (== groups_shortfall), so the audit trail on the
    Extraction Check tab could claim "assumed 0 additional vacant lots"
    while 3 phantom rows were silently appended to rentRollRows."""
    financials = _build_financials()
    # Make unitGroups already tally to stated_units (groups_shortfall=0)
    # while rentRollRows stays at 97 (rows_shortfall=3).
    financials["rentRoll"]["unitGroups"][0]["vacantCount"] = SHORTFALL
    property_info = {"units": STATED_UNITS}

    backend._ensure_rent_roll_complete(financials, property_info)

    # Rows still get imputed correctly regardless of the message bug.
    rows = financials["rentRoll"]["rentRollRows"]
    assert len(rows) == STATED_UNITS

    flags = financials.get("flags") or []
    flag = next(f for f in flags if f.get("item") == "Rent roll vs unit count")
    assert f"{SHORTFALL} additional vacant lot" in flag["issue"], flag
    assert "assumed 0" not in flag["issue"], flag

    checks = financials.get("_extractionChecks") or []
    check = next(c for c in checks if c.get("item") == "Vacant-pad imputation")
    assert f"Imputed {SHORTFALL}" in check["detail"], check
    assert "Imputed 0" not in check["detail"], check
