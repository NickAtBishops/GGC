"""Deterministic tests for the per-row vacant-lot rent imputation inside
`backend.fill_template`'s Rent Roll Input writer.

CLAUDE.md §2.3 / §5.1 are explicit: a vacant pad on the rent roll must
NEVER hit Rent Roll Input column I as $0. The seller's rent roll usually
shows vacants as blanks; if we write those blanks straight through, GPR
silently undercounts and the whole collections build (Stage 3 / §5.1) is
biased low. The imputation rule:

    vacant row with lotRent <= 0  ->  market rent for the row's
                                       canonical unit type, sourced
                                       preferentially from
                                       rentRoll.unitGroups[*].lotRent,
                                       falling back to the average of
                                       occupied per-row lotRents of the
                                       same canonical type.

The logic lives as an inline slice inside `fill_template` (around line
7006-7221 in backend.py at the time these tests were written). Rather
than spinning up the full 16-tab workbook write, this test does what
test_canonicalize_unit_type.py does: source-slice the relevant block out
of `fill_template`, exec it in an isolated namespace with `wb["Rent Roll
Input"]` mocked to a MagicMock, then inspect the recorded `ws.cell`
calls. That exercises the SAME production code path (no copy that can
drift) while keeping the test deterministic and openpyxl-free.

Three cases:
1. Vacant TOH row with lotRent=0 picks up the unitGroups TOH market rent.
2. Vacant POH row with lotRent=0 picks up the unitGroups POH market rent
   (NOT the TOH rent — per-type lookup must be honored).
3. Vacant row with lotRent already > 0 (a seller who DID fill in market
   rents on the rent roll) is written through untouched — no override.
"""
from __future__ import annotations

import inspect
import sys
import textwrap
from pathlib import Path
from unittest.mock import MagicMock

import pytest

# Make backend.py importable regardless of pytest's invocation directory.
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(REPO_ROOT))

import backend  # noqa: E402


# --------------------------------------------------------------------------- #
# Source-slice the Rent Roll Input writer out of fill_template.               #
# --------------------------------------------------------------------------- #
# The block runs from `ws = wb["Rent Roll Input"]` (the section header
# comment is right above it) through the end of the column-write loop,
# which is bounded by the `# A (Count) and K (Combined) ...` trailing
# comment and the next section header (`# ── Add Miscellaneous tab ──`).
#
# We grab everything from the ws assignment through the end of the
# `for i, unit in enumerate(...)` loop, dedent, and exec it against a
# namespace that supplies:
#   - `wb` (a MagicMock whose __getitem__ returns the worksheet mock)
#   - `_protect_formulas` (a no-op — the real one wraps openpyxl cells)
#   - `financials` (the test fixture)
SLICE_START_MARKER = '    ws = wb["Rent Roll Input"]'
# Stop right before the next ── section header so we capture the whole
# Rent Roll Input block including the write loop.
SLICE_END_MARKER   = "    # ── Add Miscellaneous tab"


def _load_rent_roll_writer_slice() -> str:
    src = inspect.getsource(backend.fill_template)
    lines = src.splitlines(keepends=True)

    start_idx = next(
        (i for i, line in enumerate(lines) if line.rstrip("\n") == SLICE_START_MARKER),
        None,
    )
    assert start_idx is not None, (
        "Could not find the `ws = wb[\"Rent Roll Input\"]` slice start "
        "inside fill_template. The writer block may have moved — update "
        "SLICE_START_MARKER in this test."
    )

    end_idx = next(
        (i for i, line in enumerate(lines[start_idx + 1:], start=start_idx + 1)
         if line.startswith(SLICE_END_MARKER)),
        None,
    )
    assert end_idx is not None, (
        "Could not find the end-of-slice marker. The Miscellaneous-tab "
        "header may have been renamed — update SLICE_END_MARKER."
    )

    snippet = "".join(lines[start_idx:end_idx])
    return textwrap.dedent(snippet)


_RENT_ROLL_WRITER_SOURCE = _load_rent_roll_writer_slice()


def _run_rent_roll_writer(financials: dict) -> MagicMock:
    """Exec the Rent Roll Input writer slice with `ws` as a MagicMock.

    Returns the worksheet mock so callers can inspect `ws.cell.call_args_list`
    and reconstruct exactly which cells were written and with what values.
    """
    ws_mock = MagicMock(name="rent_roll_input_ws")
    wb_mock = MagicMock(name="wb")
    wb_mock.__getitem__.return_value = ws_mock

    import re as _re
    ns = {
        "wb":                 wb_mock,
        "_protect_formulas":  lambda _ws: None,  # no-op stand-in
        "financials":         financials,
        # The writer slice uses regex tables for the seller-label-to-short-
        # code mapping (TOH/POH/LTO/Flourish/RV/Retail). Pass `re` in so
        # the compile calls in the slice resolve at exec time.
        "re":                 _re,
    }
    exec(_RENT_ROLL_WRITER_SOURCE, ns)
    return ws_mock


def _lot_rent_writes_by_row(ws_mock: MagicMock) -> dict[int, object]:
    """Collect the column-H (Lot Rent) writes keyed by row number.

    Layout: the new Parkwood-style Rent Roll Input has H=Lot Rent,
    I=POH Home Rents, J=LTO PMT, K=Combined. The writer calls
    `ws.cell(row=r, column=8, value=lot_rent)`.
    """
    writes: dict[int, object] = {}
    for call in ws_mock.cell.call_args_list:
        kwargs = call.kwargs
        if kwargs.get("column") == 8:
            writes[kwargs["row"]] = kwargs.get("value")
    return writes


# --------------------------------------------------------------------------- #
# Fixtures                                                                    #
# --------------------------------------------------------------------------- #
# Distinct per-type market rents make a wrong-bucket bug visible in the
# test output rather than ambiguous.
TOH_MARKET_RENT = 525.0
POH_MARKET_RENT = 825.0


def _build_financials(
    rent_roll_rows: list[dict],
    *,
    toh_market: float = TOH_MARKET_RENT,
    poh_market: float = POH_MARKET_RENT,
) -> dict:
    """Wrap the given rent roll rows in the financials shape the writer expects.

    unitGroups carries the per-canonical-type lotRent that the writer's
    market_lot_rent_by_type lookup reads first. This mirrors how a real
    methodology output is shaped — the LLM populates unitGroups summary
    statistics in addition to per-row data.
    """
    return {
        "rentRoll": {
            "totalUnits":    len(rent_roll_rows),
            "occupiedUnits": sum(
                1 for r in rent_roll_rows
                if (r.get("status") or "Occupied").lower() != "vacant"
            ),
            "vacantUnits":   sum(
                1 for r in rent_roll_rows
                if (r.get("status") or "Occupied").lower() == "vacant"
            ),
            "unitGroups": [
                {
                    "unitType":      "TOH MH Site",
                    "occupiedCount": 1,
                    "vacantCount":   0,
                    "lotRent":       toh_market,
                },
                {
                    "unitType":      "POH-Infilled units",
                    "occupiedCount": 1,
                    "vacantCount":   0,
                    "lotRent":       poh_market,
                },
            ],
            "rentRollRows": rent_roll_rows,
        }
    }


def _occupied(unit_id: str, unit_type: str, lot_rent: float) -> dict:
    return {
        "tenantName": f"Tenant {unit_id}",
        "unitId":     unit_id,
        "unitType":   unit_type,
        "status":     "Occupied",
        "lotRent":    lot_rent,
        "homeRent":   0,
        "lcPayment":  0,
        "moveInDate": "2020-01-01",
    }


def _vacant(unit_id: str, unit_type: str, lot_rent: float = 0) -> dict:
    return {
        "tenantName": "",
        "unitId":     unit_id,
        "unitType":   unit_type,
        "status":     "Vacant",
        "lotRent":    lot_rent,
        "homeRent":   0,
        "lcPayment":  0,
        "moveInDate": "",
    }


# --------------------------------------------------------------------------- #
# Tests                                                                       #
# --------------------------------------------------------------------------- #
def test_vacant_toh_row_imputes_toh_market_rent_from_unit_groups():
    """A vacant TOH row with lotRent=0 must be written at the TOH
    unitGroups market rent, NOT $0."""
    rows = [
        _occupied("1", "TOH MH Site", 500.0),    # row 3 in the worksheet
        _vacant("2",   "TOH MH Site"),           # row 4 — the imputation target
    ]
    financials = _build_financials(rows)

    ws_mock = _run_rent_roll_writer(financials)

    lot_rent_writes = _lot_rent_writes_by_row(ws_mock)
    assert 4 in lot_rent_writes, (
        f"Expected a lot-rent write at row 4 (the vacant TOH row); "
        f"saw writes at rows {sorted(lot_rent_writes.keys())}."
    )
    assert lot_rent_writes[4] == pytest.approx(TOH_MARKET_RENT), (
        f"Vacant TOH row should pick up TOH unitGroups market rent "
        f"({TOH_MARKET_RENT}); got {lot_rent_writes[4]!r}. A $0 here is "
        "the GPR-undercount bug CLAUDE.md §2.3 forbids."
    )
    # Sanity: the occupied row still reflects its own contracted rent.
    assert lot_rent_writes[3] == pytest.approx(500.0), (
        f"Occupied row should pass through its contracted lot rent of 500; "
        f"got {lot_rent_writes[3]!r}. The imputation pass should NEVER "
        "touch occupied rows."
    )


def test_per_type_market_rent_lookup_does_not_cross_pollinate():
    """Imputation is per canonical unit type. A vacant POH row must get
    the POH market rent, not the TOH market rent — even when TOH and POH
    appear in the same rent roll."""
    rows = [
        _occupied("1", "TOH MH Site",        500.0),   # row 3
        _occupied("2", "POH-Infilled units", 800.0),   # row 4
        _vacant("3",   "POH-Infilled units"),          # row 5 — imputes POH
        _vacant("4",   "TOH MH Site"),                 # row 6 — imputes TOH
    ]
    financials = _build_financials(rows)

    ws_mock = _run_rent_roll_writer(financials)
    lot_rent_writes = _lot_rent_writes_by_row(ws_mock)

    assert lot_rent_writes[5] == pytest.approx(POH_MARKET_RENT), (
        f"Vacant POH row (worksheet row 5) should pick up POH market rent "
        f"({POH_MARKET_RENT}); got {lot_rent_writes[5]!r}. A TOH value here "
        "would mean the per-type lookup is broken and POH GPR is mis-priced."
    )
    assert lot_rent_writes[6] == pytest.approx(TOH_MARKET_RENT), (
        f"Vacant TOH row (worksheet row 6) should pick up TOH market rent "
        f"({TOH_MARKET_RENT}); got {lot_rent_writes[6]!r}."
    )


def test_vacant_row_with_explicit_rent_is_passed_through_unchanged():
    """A seller who DID fill in a market rent on the vacant row must
    win — imputation only fires when lotRent <= 0, per the §2.3 rule."""
    explicit_vacant_rent = 612.34
    rows = [
        _occupied("1", "TOH MH Site", 500.0),                         # row 3
        _vacant("2",   "TOH MH Site", lot_rent=explicit_vacant_rent), # row 4
    ]
    financials = _build_financials(rows)

    ws_mock = _run_rent_roll_writer(financials)
    lot_rent_writes = _lot_rent_writes_by_row(ws_mock)

    assert lot_rent_writes[4] == pytest.approx(explicit_vacant_rent), (
        f"Vacant row with an explicit lotRent ({explicit_vacant_rent}) "
        f"should pass through unchanged; got {lot_rent_writes[4]!r}. "
        "Imputation is a backstop, not an override."
    )


def test_status_column_records_vacant_for_imputed_rows():
    """Defense-in-depth: even after imputation rewrites lotRent, the
    row's status cell (column E in the new Parkwood layout) must still
    read 'Vacant' so the workbook's Unit Mix Summary occupancy COUNTIFS
    stay honest. Layout: B=Lot#, C=Lot Type, D=Type formula, E=Status,
    F=Tenant, G=Move in, H=Lot Rent."""
    rows = [
        _occupied("1", "TOH MH Site", 500.0),    # row 3
        _vacant("2",   "TOH MH Site"),           # row 4
    ]
    financials = _build_financials(rows)

    ws_mock = _run_rent_roll_writer(financials)

    status_writes = {
        call.kwargs["row"]: call.kwargs.get("value")
        for call in ws_mock.cell.call_args_list
        if call.kwargs.get("column") == 5
    }
    assert status_writes.get(4) == "Vacant", (
        f"Expected column E / row 4 to be 'Vacant'; got {status_writes.get(4)!r}. "
        "Imputation must not flip a vacant row's status to occupied — that "
        "would corrupt the occupancy rate (N8) and every downstream calc."
    )
