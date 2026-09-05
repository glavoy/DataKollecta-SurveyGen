"""Fixtures for testing a generated package, not for generating one.

Nothing in here is imported by the generator. It exists so a real dictionary
can be turned into the set of variants that sweep the `auto_start_repeat` x
`repeat_enforce_count` matrix, which no real dictionary covers on its own.

Present so `python -m test_kit.make_variants` puts the repo root on `sys.path`
rather than `test_kit/`, which is what lets this module import `crf_reader`
and `processor` the same way `tests/` does. Same reason `tests/__init__.py`
exists.
"""
