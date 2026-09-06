# To Do

- **Validate `FieldName` and the other identifiers against the rule the README already
  states.** The app now refuses to install a package whose `FieldName`, `<responses
  source="database">` `table`/`column`, or `crfs` identifier cell is not a plain identifier
  (letter or underscore, then letters, digits and underscores) -- see
  `SurveyTableSchema.validateIdentifier` in the app repo. This generator does not check any
  of it, so a dictionary with `first name` in it builds a clean package that fails at
  install, in the field, rather than at authoring time where the fix is. The README's
  FieldName section documents the rule and the examples are already written; nothing
  enforces them. `skip_parser.py`'s `_IDENTIFIER` is the same pattern, so the assumption is
  already implicit in the skip grammar.
