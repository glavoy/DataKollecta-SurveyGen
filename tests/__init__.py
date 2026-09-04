"""Marks `tests` as a package so the suite's own cross-imports are declared.

Several test modules share fixtures -- `from tests.test_dd_validations import
HEADERS`, `from tests.test_system_fields_xml import generate` -- which worked
only because the repo root happens to be on `sys.path` when the suite is run
from there, with `tests` resolving as an implicit namespace package.

Making it explicit costs nothing and buys one thing: both discovery forms now
work. Measured on a clean `git archive` of the tree, before and after:

    python -m unittest discover -s tests        without: OK    with: OK
    python -m unittest discover -s tests -t .   without: FAIL  with: OK

The `-t .` failure is "Start directory is not importable" -- that form
requires `tests` to be a real package. Nothing regressed, so the file is worth
having.
"""
