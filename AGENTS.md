# Repository Guidelines

## Project Structure & Module Organization
Source modules live at the repository root: `mail_grep.py` is the CLI entry point, while helpers such as `mail_folder.py`, `mail_message.py`, `mail_profile.py`, and `search_pattern.py` encapsulate filesystem access, parsing, and matching responsibilities. Reporting utilities (`hit_line.py`, `hit_report.py`, `smart_logging.py`) shape the CSV/XLSX output written to `results/`. Reference material and diagrams stay under `docs/`, with visual assets in `docs/assets/`. Temporary exports and scratch files should never be committed.

## Build, Test, and Development Commands
Create an isolated environment and install the tool with `python -m pip install -e .`. Run the CLI locally via `python -m mail_grep "<regex>" -s /path/to/Mail`. Use `python -m mail_grep --help` to inspect flags, and set `MAIL_GREP_LOG=DEBUG` before execution when you need verbose SmartLogging output. Regenerate sample reports with `python -m mail_grep "Invoice" -o results/invoice_demo.csv` to confirm formatting.

## Coding Style & Naming Conventions
Follow Python 3.10+ conventions: four-space indentation, trailing newline, and type hints for public APIs. Keep modules focused and prefer small, single-responsibility classes similar to `MailMessage` and `HitReport`. Use `snake_case` for functions and variables, `PascalCase` for classes, and `SCREAMING_SNAKE_CASE` for constants. Before pushing, run `python -m compileall .` or your formatter of choice (Black is recommended even though it is not pinned) to catch syntax errors.

## Testing Guidelines
The project currently relies on manual verification. When adding behavior, create targeted unit tests under a new `tests/` package using `pytest` (`python -m pytest`) or the standard library `unittest`. Mock filesystem-dependent calls (`MailFolder.mail_paths`) and supply fixture `.emlx` samples stored in `tests/data/`. Always run the CLI against a small Mail subset to verify CSV/XLSX output and date ordering after test execution.

## Commit & Pull Request Guidelines
History favors concise, lowercase commit subjects (`git log` shows entries like `update` and `nits`); continue using short present-tense summaries and include detail in the body when needed. Each pull request should describe the motivation, key changes, and manual verification steps, and link the relevant issue or ticket. Attach before/after log snippets or sample CSV rows for features that affect reporting, and call out any new dependencies or migration steps explicitly.
