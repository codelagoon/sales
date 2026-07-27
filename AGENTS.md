# AGENTS.md

## Project

The product is a single self-contained, macro-enabled Excel workbook,
`Assigned Late Schedule.xlsm`, generated from source. There is no server or web
app. See `README.md` for the full description and commands.

**Runtime target:** Microsoft Excel for Windows only. Do not attempt Mac Excel
compatibility unless explicitly requested — the VBA engine depends on
`Scripting.Dictionary` and other Windows COM features.

- `build_workbook.py` — builds the workbook layout with XlsxWriter and embeds
  `assets/vbaProject.bin`.
- `vba/AssignedLateSchedule.bas` — the VBA (date generation, fairness optimizer,
  history/audit, transparency summary, safeguards, verification).
- `tools/compile_vba.py` — compiles the `.bas` into `assets/vbaProject.bin`
  using pure-Python `pyOpenVBA` (no Excel needed).
- `tests/schedule_engine.py` — a faithful Python port of the VBA engine.
- `tests/test_schedule.py` — the test suite.

## Cursor Cloud specific instructions

- Build/test toolchain is pure Python (XlsxWriter, pyOpenVBA, oletools); no
  Microsoft Excel or LibreOffice is required to build or test.
- Standard commands (see `README.md`): `python tools/compile_vba.py` then
  `python build_workbook.py`; tests via `python tests/test_schedule.py`;
  VBA name-collision guard via `python tools/check_vba_names.py` (run before
  compile — case-insensitive Dim/Function name clashes cause runtime error 91).
- The VBA cannot execute headlessly here (no Excel). The Python port in
  `tests/schedule_engine.py` mirrors the VBA exactly and is the way to verify
  scheduling/fairness logic. **Keep it in sync with `vba/AssignedLateSchedule.bas`**:
  change the algorithm in both and re-run the tests.
- `assets/vbaProject.bin` must contain the `AssignedLateSchedule` module. After
  editing the `.bas`, always re-run `python tools/compile_vba.py` or the four
  buttons will have no code behind them. Verify with:
  `python -c "from oletools.olevba import VBA_Parser; print([n for _,_,n,_ in VBA_Parser('Assigned Late Schedule.xlsm').extract_macros()])"`.
  Excel recompiles the VBA p-code from the embedded source on first open.
- Git auth gotcha: the global `~/.gitconfig` (at `/home/ubuntu/.gitconfig`)
  injects the push token via `url.*.insteadOf`. Keep `HOME=/home/ubuntu` for any
  git operation. Do **not** `export HOME=/tmp` globally (it hides that config and
  breaks `git push`); if a tool needs a different HOME (e.g. LibreOffice), set it
  inline for that one command only, e.g. `HOME=/tmp soffice ...`.
- Optional visual rendering only: `libreoffice-calc` + `poppler-utils`
  (`pdftoppm`) can render the sheet to PDF/PNG. LibreOffice omits the form-control
  buttons in PDF export and does not run the VBA, so use it for layout previews
  only, not macro testing. These are not needed for build or tests.
