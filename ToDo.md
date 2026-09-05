# To Do

## Found during the M5 decomposition - each its own commit

- **`_to_str = staticmethod(to_str)` in `excel_reader.py` has no call sites.** The other three
  `cell_text` aliases are used (`_split_lines` in four places, `_get_cell_trim` and
  `_get_cell_raw` in their wrappers); this one is dead and predates the reader split. Delete it,
  or find what was meant to use it. `to_str` itself is used - by `cell_trim`/`cell_raw`
  internally and by `tests/test_cell_datetime.py` - so only the alias goes.
