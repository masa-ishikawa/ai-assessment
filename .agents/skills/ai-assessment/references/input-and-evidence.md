# Generator content handoff

## Project and input routes

Project root: use the current Codex workspace or walk upward until both `generate_assessment.py` and `.agents/skills/ai-assessment/SKILL.md` exist. The skill may also be installed in a user-level folder, so do not derive the project path from this file.
Read applicable project instructions and actual CLI interfaces before executing. Locate the project if moved; do not duplicate its parser or recreate its content pipeline inside the skill.

Prefer an existing `ai-assess/assessment-v2` output with `input`, `assessment`, `research`, `rendering` and `provenance`, plus the corresponding research log and latest PPTX. JSON supplies complete content; the source PPTX governs presentation coverage, page grouping and body composition. If only a PPTX exists, inspect its visible slides, tables and notes, and look for its JSON/log before reconstructing any missing content.

When new content is required, use the unchanged generator:

```sh
python scripts/run_assessment.py generate <input.xlsx-or-txt> --json-only --output-json <new-assessment.json>
python scripts/run_assessment.py generate --validate-json <new-assessment.json>
```

Replace placeholders with actual quoted paths and confirm current CLI behavior. Review the resulting content and research log even if validation passes. Use unchanged `--from-json <assessment.json> --output <private-reference.pptx>` when a target composition reference is needed. Restyle that reference through the isolated Oracle authoring path. Freeze means do not edit the implementation; executing its content and rendering paths is permitted.

- Fixed ISV Excel uses the existing extraction/normalization modules. The read-only `--extract-input-json` route can inspect answers without generating content; check `missing_required`, not just exit code.
- Supported text and legacy `サービス入力` workbooks use `ai_assess_runtime.source_input.load_source_text`. The implementation supports `.xlsx` and text, not binary `.xls`.
- URL-only requests do not require creating Excel. Research the official identity, product, supplied intent and source URLs into a UTF-8 text input, mark unknowns, then feed it through the generator. Do not independently select the final portfolio or author business effects as a replacement for generator output.
- Do not write answers into the input workbook unless requested. Preserve source files and their hashes.

If the content pipeline is unavailable or fails, retain available outputs, report the exact limitation and continue independent inspection/layout planning. Codex-authored replacement assessment content requires an explicit change of approach, not an automatic fallback.

## Evidence and permitted editorial changes

Preserve the generator's company/service identity, user priorities, selected use cases, strategic argument, business effects, PoC logic, technical proposal and support plan. Faithful shortening, line breaks and visual grouping are layout edits. Changing selected themes, causal arguments, financial assumptions, success thresholds or substantive section coverage is content revision and must be explicitly requested or identified as a proposed correction.

Review external outcomes, company targets, generated assumptions and proposed acceptance criteria separately. A `verified_external` tag alone does not validate a mismatched excerpt or metric. Check URLs, definitions, periods, units, price date/SKU, arithmetic and exclusions. Flag unsupported or inconsistent claims; do not reproduce them as verified facts or silently fabricate replacements. Record correction/omission reasons and surface material discrepancies.

Additional research may verify a claim or fill an explicitly requested gap. It must not silently replace the existing generator's research and reasoning process. Unagreed estimates remain unagreed; do not promote them into customer targets or guarantees.

## Authoring handoff and records

Keep a separate `ai-assessment/outline-v1` JSON with:

- Original input, generator JSON/PPTX and research paths/hashes.
- Company/service scope, full source use cases and selected IDs.
- Source section keys and JSON paths mapped to output slide numbers.
- Original substantive text, faithful display text, notes and source IDs.
- Selected Oracle base compositions and layout decisions.
- Explicit discrepancies, proposed corrections and any approved changes.

Do not overwrite or relabel generator JSON as the authoring outline. Preserve the substantive content coverage and page grouping of the source PPTX. Full JSON supports traceability and detail; do not turn each nested field into additional slides. Notes may retain supporting detail and citations. Keep diagnostics, reproduction commands and QA logs separate. Verify frozen source hashes when reporting that implementation was unchanged.
