This repository contains a configuration-driven stock analysis agent.

Project conventions:
- Keep AI-based and rule-based recommendations independent so the output can compare both tracks fairly.
- Treat `config.json` as the source of truth for stock symbols, output settings, history settings, and email settings.
- Preserve the current Excel-first workflow unless the task explicitly asks to change the export format.
- When adding metrics or ranking logic, update both the implementation and the user-facing documentation.
- If historical comparison is available, prefer using archived run snapshots rather than inventing week-over-week changes.
- If a metric is unavailable from the upstream source, leave a clear fallback value instead of guessing.

When working in this repo:
- Prefer small, root-cause changes in `Stock_Agent.py`.
- Validate with `python -m py_compile .\Stock_Agent.py` before wider testing.
- Keep generated artifacts and credentials out of committed instructions.