"""
Regression tests for `_infer_operative_period` (backend.py).

CLAUDE.md §2.2/§7 describe "a dedicated period-identification heuristic
(regex on column headers + 12-month count)" as an existing mitigation for
the multi-period problem. An accuracy audit of the Las Brisas deal (a real
document with genuinely multiple period columns) found this heuristic did
not actually exist — period selection was 100% delegated to LLM judgment
with no deterministic cross-check. `_infer_operative_period` closes that
gap: it parses the LLM's own self-reported `candidatePeriodsSeen` list for
T-12/date-range headers and independently identifies the most recent
full-12-month candidate, so `verify_extraction` can raise a WARN when it
disagrees with the model's chosen `periodUsed`. It must NEVER silently
override the model's choice — only flag disagreement.

Run: pytest tests/test_period_inference.py -v
"""
from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_ROOT))

import backend  # noqa: E402


def test_las_brisas_real_headers_pick_most_recent_t12():
    """The exact 3 period-column headers from the real Las Brisas T-12
    file. "T-12 Ended 5/23" is the correct trailing-12 column (verified by
    the arithmetic identity N - stub + partial = P against the real file);
    "T-12 Ended 9/22" is stale; "Oct 2022- May 2023" is an 8-month partial
    that must be excluded from consideration entirely.
    """
    candidates = ["T-12 Ended 9/22", "Oct 2022- May 2023", "T-12 Ended 5/23"]
    guess, reason = backend._infer_operative_period(candidates, "T-12 Ended 5/23")
    assert guess == "T-12 Ended 5/23"
    assert "matches periodUsed" in reason


def test_disagreement_is_detected_not_silently_fixed():
    """If the model picks the stale column, the function must say so —
    and the caller (verify_extraction) must only WARN, never rewrite
    periodUsed itself.
    """
    candidates = ["T-12 Ended 9/22", "Oct 2022- May 2023", "T-12 Ended 5/23"]
    guess, reason = backend._infer_operative_period(candidates, "T-12 Ended 9/22")
    assert guess == "T-12 Ended 5/23"
    assert "matches periodUsed" not in reason
    assert "T-12 Ended 9/22" in reason


def test_partial_period_excluded_from_candidacy():
    """An 8-month range must never be treated as a full-T12 candidate,
    even if it's the only thing that looks date-like.
    """
    guess, reason = backend._infer_operative_period(["Oct 2022- May 2023"], None)
    assert guess is None
    assert "no candidate" in reason


def test_no_recognizable_format_returns_none_not_error():
    """Unrecognized header formats must return (None, reason) — 'no
    opinion,' not a false disagreement flag. The design must not punish
    documents whose headers this regex doesn't happen to recognize.
    """
    guess, reason = backend._infer_operative_period(
        ["Column A", "Prior Year", "Current Year"], "Current Year")
    assert guess is None


def test_empty_candidates_returns_none():
    guess, reason = backend._infer_operative_period([], "T-12 Ended 5/23")
    assert guess is None


def test_ttm_and_trailing_12_synonyms_recognized():
    """Real sellers use several synonyms for the same trailing-12 concept."""
    guess, _ = backend._infer_operative_period(["TTM Ended 3/24"], "TTM Ended 3/24")
    assert guess == "TTM Ended 3/24"
    guess2, _ = backend._infer_operative_period(["Trailing 12 Ended 3/24"], None)
    assert guess2 == "Trailing 12 Ended 3/24"


def test_two_digit_year_normalized_correctly():
    """'5/23' must resolve to 2023, not 1923 or a raw '23'."""
    guess, _ = backend._infer_operative_period(["T-12 Ended 5/23"], "T-12 Ended 5/23")
    assert guess == "T-12 Ended 5/23"
    # Confirm ordering logic actually parses the year, not just string-sorts
    # ("5/9" would string-sort after "5/23" if year parsing were broken).
    guess2, reason2 = backend._infer_operative_period(
        ["T-12 Ended 5/9", "T-12 Ended 5/23"], "T-12 Ended 5/9")
    assert guess2 == "T-12 Ended 5/23"
    assert "matches periodUsed" not in reason2


def test_full_date_range_recognized_as_t12_when_exactly_12_months():
    """A "Mon YYYY - Mon YYYY" header spanning exactly 12 months (not the
    T-12-labeled idiom) must also be recognized as a valid full-period
    candidate — the range math, not just the "T-12" keyword, decides.
    """
    guess, reason = backend._infer_operative_period(
        ["Jan 2023 - Dec 2023"], "Jan 2023 - Dec 2023")
    assert guess == "Jan 2023 - Dec 2023"
    assert "matches periodUsed" in reason


def test_verify_extraction_integration_only_warns_on_real_disagreement():
    """End-to-end through verify_extraction: silent when the model agrees,
    a precise WARN (never a silent rewrite) when it doesn't.
    """
    candidates = ["T-12 Ended 9/22", "Oct 2022- May 2023", "T-12 Ended 5/23"]

    correct = {"reportingPeriod": {
        "periodUsed": "T-12 Ended 5/23", "dateRange": "Jun 2022 - May 2023",
        "monthsCovered": 12, "candidatePeriodsSeen": candidates, "notes": "",
    }}
    checks = backend.verify_extraction(correct, {})
    cross_check = [c for c in checks if c["item"] == "Period selection cross-check"]
    assert not cross_check, "must not fire when the model already picked correctly"

    wrong = {"reportingPeriod": {
        "periodUsed": "T-12 Ended 9/22", "dateRange": "Oct 2021 - Sep 2022",
        "monthsCovered": 12, "candidatePeriodsSeen": candidates, "notes": "",
    }}
    checks2 = backend.verify_extraction(wrong, {})
    cross_check2 = [c for c in checks2 if c["item"] == "Period selection cross-check"]
    assert len(cross_check2) == 1
    assert cross_check2[0]["status"] == "warn"
    assert "T-12 Ended 5/23" in cross_check2[0]["detail"]
    # periodUsed itself must be untouched — this is a flag, not a rewrite.
    assert wrong["reportingPeriod"]["periodUsed"] == "T-12 Ended 9/22"
