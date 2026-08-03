"""Deterministic tests for `_canonicalize_unit_type` (backend.py).

The function lives as a closure inside `backend.fill_template` (around
line 7032). It maps a seller's raw unit-type label to one of the
canonical Parkwood-era unit-type strings the 16-tab template's Unit Mix
Summary and Rent Roll Input tabs expect:

  - "TOH MH Site"        — plain tenant-owned MH lot
  - "POH-Infilled units" — park-owned home (incl. title-issue infills)
  - "LTO MH Site"        — Land Contract / Lease-to-Own (TOH-LC variants)
  - "Flourish MH Site"   — Flourish / Bennetts / trustee sub-brand TOH
  - "Long term RV Site"  — RV (annual / long-term)
  - "Retail/Commercial"  — retail / storage / boat / garage

These canonicals drive the lot-rent-vs-home-rent bifurcation (§5.2) and
the vacant-pad market-rent lookup (§2.3 / §5.1), so a drift in this
mapping silently mis-routes whole categories of GPR. Pin it.

Because the function is a nested closure (not a module-level symbol),
this test extracts its source via `inspect.getsource` on the parent
`fill_template`, slices out the `def _canonicalize_unit_type` block,
and `exec`s it into a private namespace. This keeps the test
exercising the SAME code path that `fill_template` runs in production —
no separate copy that can drift out of sync.
"""
from __future__ import annotations

import inspect
import sys
import textwrap
from pathlib import Path

import pytest

# Make backend.py importable regardless of pytest's invocation directory.
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(REPO_ROOT))

import backend  # noqa: E402


# --------------------------------------------------------------------------- #
# Extract the nested `_canonicalize_unit_type` closure from fill_template.    #
# --------------------------------------------------------------------------- #
def _load_canonicalize_unit_type():
    """Slice the nested def out of fill_template and exec it in isolation.

    The function only depends on `re` (imported locally inside the def
    itself as `import re as _re`), so it has no other free variables
    from the enclosing fill_template scope. That makes safe extraction
    by source-slicing possible.
    """
    src = inspect.getsource(backend.fill_template)

    marker = "def _canonicalize_unit_type(raw, home_rent=0):"
    lines = src.splitlines(keepends=True)
    start_idx = next(
        (i for i, line in enumerate(lines) if marker in line), None
    )
    assert start_idx is not None, (
        "Could not find _canonicalize_unit_type inside fill_template. "
        "The function may have been renamed or moved — update this test."
    )

    # Indent level of the `def _canonicalize_unit_type` line itself.
    # The function body is indented one more level than this, and the
    # def ends when we see another non-blank line at this same indent
    # (a peer statement in fill_template).
    def_line = lines[start_idx]
    def_indent = len(def_line) - len(def_line.lstrip(" "))

    body_lines = [def_line]
    for line in lines[start_idx + 1:]:
        stripped_nl = line.lstrip(" ")
        # Blank or whitespace-only line — keep it as part of the body.
        if not stripped_nl.strip():
            body_lines.append(line)
            continue
        line_indent = len(line) - len(stripped_nl)
        # First non-blank line back at the def's indent (or shallower)
        # is the next peer statement — the def is done.
        if line_indent <= def_indent:
            break
        body_lines.append(line)

    # Dedent the captured snippet so the def starts at column 0.
    snippet = textwrap.dedent("".join(body_lines))
    ns: dict = {}
    exec(snippet, ns)
    fn = ns.get("_canonicalize_unit_type")
    assert callable(fn), "Extraction did not produce a callable function."
    return fn


_canonicalize_unit_type = _load_canonicalize_unit_type()


# --------------------------------------------------------------------------- #
# Cases — one assertion per row. Parametrize so a single regression is        #
# obvious from the pytest output rather than buried in a multi-assert test.   #
# --------------------------------------------------------------------------- #
CASES = [
    # Plain TOH — tenant-owned MH lot.
    ("TOH", "TOH MH Site"),
    # POH — park-owned home, routes to the Infilled-units bucket.
    ("POH", "POH-Infilled units"),
    # TOH-LC — Land Contract variant, economically an LTO, NOT plain TOH.
    # This is the OUTLINE root-cause-#1 case: must NOT collapse into TOH.
    ("TOH-LC", "LTO MH Site"),
    # "Land Contract" spelled out — same LTO bucket.
    ("Land Contract", "LTO MH Site"),
    # Bare "LTO" abbreviation.
    ("LTO", "LTO MH Site"),
    # "Lease to Own" spelled out — same LTO bucket.
    ("Lease to Own", "LTO MH Site"),
    # Flourish sub-brand — distinct from plain TOH, has its own canonical.
    ("Flourish", "Flourish MH Site"),
    # Bennett(s) — Flourish-family financing, routes to Flourish bucket.
    ("Bennett", "Flourish MH Site"),
    # RV Annual — long-term RV site, distinct from MH lots.
    ("RV Annual", "Long term RV Site"),
    # Retail storefront — Retail/Commercial bucket.
    ("Retail", "Retail/Commercial"),
]


@pytest.mark.parametrize("raw,expected", CASES, ids=[c[0] for c in CASES])
def test_canonicalize_unit_type(raw: str, expected: str) -> None:
    """Each raw seller label must map to its canonical Parkwood string."""
    actual = _canonicalize_unit_type(raw)
    assert actual == expected, (
        f"_canonicalize_unit_type({raw!r}) returned {actual!r}, "
        f"expected {expected!r}. A drift here silently mis-routes GPR "
        f"between TOH / LTO / Flourish / POH / RV / Retail buckets."
    )


# --------------------------------------------------------------------------- #
# home_rent fallback — opaque/coded unit-type labels (e.g. Blue Island's      #
# bare "Type 1"/"Type 2"/"Type 4") that match no keyword above must NOT       #
# silently collapse into the TOH catch-all when the row is actually charging  #
# home rent. This is a seller-agnostic signal, not a hardcoded mapping for    #
# any one deal's coding convention (CLAUDE.md forbids hardcoding one         #
# seller's layout into core logic) — it only fires when nothing else does.   #
# --------------------------------------------------------------------------- #
HOME_RENT_FALLBACK_CASES = [
    # Opaque numeric code, no home rent charged -> default TOH, unchanged.
    ("Type 1", 0, "TOH MH Site"),
    # Same opaque code, but this row DOES charge home rent -> POH signal
    # wins even though "Type 2" itself matches no keyword.
    ("Type 2", 225, "POH-Infilled units"),
    # A keyword match (RV) must still win over the home_rent signal even
    # if home_rent is nonzero for some other reason (data noise) — the
    # fallback only applies when NO keyword matched.
    ("RV", 100, "Long term RV Site"),
    # Blank/missing type string, no home rent -> default TOH.
    ("", 0, "TOH MH Site"),
    # Blank/missing type string, WITH home rent -> POH.
    ("", 500, "POH-Infilled units"),
]


@pytest.mark.parametrize(
    "raw,home_rent,expected",
    HOME_RENT_FALLBACK_CASES,
    ids=[f"{c[0] or '<blank>'}-homeRent{c[1]}" for c in HOME_RENT_FALLBACK_CASES],
)
def test_canonicalize_unit_type_home_rent_fallback(
    raw: str, home_rent: float, expected: str
) -> None:
    actual = _canonicalize_unit_type(raw, home_rent)
    assert actual == expected, (
        f"_canonicalize_unit_type({raw!r}, home_rent={home_rent!r}) "
        f"returned {actual!r}, expected {expected!r}. An opaque unit-type "
        f"code with a real home-rent charge must route to POH, not "
        f"silently collapse to TOH (the Blue Island failure mode)."
    )
