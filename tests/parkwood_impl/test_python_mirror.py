"""Deterministic tests for the Python-side IRR/EM/CoC mirror helpers.

CONTEXT (CLAUDE.md §0 / backend.py lines 6448-6576)
---------------------------------------------------
`compute_underwriting_summary` (backend.py ~6579) is the Python-side
mirror of the 16-tab workbook's headline returns numbers (IRR, equity
multiple, cash-on-cash). It exists because openpyxl never writes cached
formula values, so the result panel needs a numeric banner that does NOT
depend on Excel round-tripping. The three private helpers it leans on
must be deterministic and dependency-free (no numpy_financial):

  - `_compute_irr(cashflows)` — Newton's method with a bisection
    fallback. Returns the IRR as a decimal (0.10 = 10%) or None when the
    series is degenerate (no sign change, no real root).
  - `_compute_equity_multiple(cashflows)` — sum(positive) / abs(sum(negative)).
  - `_compute_cash_on_cash(cashflows_after_debt, equity_contributed)` —
    average annual yield as a decimal.

WHAT THESE TESTS PIN
--------------------
1. IRR of [-100, 110] solves to exactly 10% (the textbook one-period case
   the Newton solver should nail on the first or second iteration).
2. IRR of [-100, 50, 60, 70] is strictly positive (the series clearly
   makes money; the exact value is verified at ~33.87% to defend against
   an accidental sign flip or solver regression).
3. IRR of [-100, 0, 0, 0] returns None — there is no positive cashflow,
   so IRR is undefined. The degenerate-input guard at line 6467 enforces
   this.
4. EM of [-100, 50, 60, 70] = 180 / 100 = 1.8.
5. CoC of [50, 60, 70] against an equity base of 100 averages to 60%
   (the three annual yields are 50%, 60%, 70%).

DESIGN NOTES
------------
- The helpers are module-level functions in `backend.py`, so we import
  them directly (same pattern as `test_monthly_fanout.py`).
- `math.isclose` with `abs_tol=1e-9` is used for float comparisons per
  the task spec. The Newton solver's `tol=1e-7` default is the natural
  noise floor, so 1e-9 is conservative but still well inside the
  solver's promised precision for the cases we exercise.
- We do NOT exercise `compute_underwriting_summary` itself here — that
  function pulls together the 10-year NOI ladder and debt-service
  schedule (backend.py ~6664-6786) and is best tested via the
  end-to-end fixture in `test_pipeline.py`. These tests pin the
  arithmetic primitives only.

Run:  pytest tests/parkwood_impl/test_python_mirror.py -v
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(REPO_ROOT))

import backend  # noqa: E402

_compute_irr = backend._compute_irr
_compute_equity_multiple = backend._compute_equity_multiple
_compute_cash_on_cash = backend._compute_cash_on_cash


# Float-comparison tolerance. Newton's solver runs to tol=1e-7 internally
# (backend.py:6448), so 1e-9 absolute is conservative and still well
# inside the solver's promised precision for our cases.
ABS_TOL = 1e-9


# ─────────────────────────────────────────────────────────────────────── #
# IRR                                                                     #
# ─────────────────────────────────────────────────────────────────────── #
def test_irr_one_period_solves_to_ten_percent():
    """The textbook one-period case: -100 today, +110 next year => 10%.

    This is the sanity case Newton's method should nail on the first or
    second iteration. If this drifts, the solver itself is broken.
    """
    irr = _compute_irr([-100, 110])
    assert irr is not None, "IRR of [-100, 110] should be defined"
    assert math.isclose(irr, 0.10, abs_tol=ABS_TOL), (
        f"Expected IRR == 0.10 (10%), got {irr!r}"
    )


def test_irr_multi_period_positive_when_series_makes_money():
    """[-100, 50, 60, 70] returns $180 against $100 in — strongly
    positive IRR. We pin both the sign (must be > 0) and a tighter
    expected value (~33.87%) so a sign flip or solver regression is
    caught by more than just the > 0 check.
    """
    irr = _compute_irr([-100, 50, 60, 70])
    assert irr is not None, "IRR of [-100, 50, 60, 70] should be defined"
    assert irr > 0, f"Expected positive IRR for a money-making series, got {irr!r}"
    # Reference value computed independently (NPV = 0 at r ≈ 0.33875).
    # Tolerance is loose enough that solver-internals tweaks won't flap
    # the test, but tight enough that a real regression bites.
    assert math.isclose(irr, 0.33875, abs_tol=1e-4), (
        f"Expected IRR ~ 0.33875, got {irr!r}"
    )


def test_irr_degenerate_no_positive_cashflows_returns_none():
    """[-100, 0, 0, 0] has no positive cashflow, so IRR is undefined.

    The guard at backend.py:6467 (`not any(c > 0 for c in cfs)`) should
    catch this before the solver even runs. Pins that we return None
    (NOT a negative number, NOT zero, NOT a silently-coerced 0.0) —
    the result panel renders "—" on None, which is the desired UX.
    """
    irr = _compute_irr([-100, 0, 0, 0])
    assert irr is None, (
        f"Expected None for a series with no positive cashflows, got {irr!r}"
    )


# ─────────────────────────────────────────────────────────────────────── #
# Equity Multiple                                                         #
# ─────────────────────────────────────────────────────────────────────── #
def test_equity_multiple_simple_case():
    """EM = sum(positive) / abs(sum(negative)) = 180 / 100 = 1.8.

    Mirrors the analyst's D27/D28 waterfall formula (total LP
    distributions / total LP contributions), per backend.py:6527.
    """
    em = _compute_equity_multiple([-100, 50, 60, 70])
    assert em is not None, "EM of [-100, 50, 60, 70] should be defined"
    assert math.isclose(em, 1.8, abs_tol=ABS_TOL), (
        f"Expected EM == 1.8 (= 180/100), got {em!r}"
    )


# ─────────────────────────────────────────────────────────────────────── #
# Cash on Cash                                                            #
# ─────────────────────────────────────────────────────────────────────── #
def test_cash_on_cash_average_yield():
    """CoC = mean(annual_cf / equity_base).

    For [50, 60, 70] against equity=100 the per-year yields are
    50%, 60%, 70%, averaging to exactly 60%. This pins both the formula
    (average, not IRR) and the decimal-not-percent convention
    (returns 0.60, not 60.0).
    """
    coc = _compute_cash_on_cash([50, 60, 70], 100)
    assert coc is not None, "CoC of [50, 60, 70] vs equity=100 should be defined"
    assert math.isclose(coc, 0.60, abs_tol=ABS_TOL), (
        f"Expected CoC == 0.60 (60%), got {coc!r}"
    )


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
