# AI Assessment — Project Instructions

## Single entry point

The user-facing entry point is the repository-local `ai-assessment` skill at `.agents/skills/ai-assessment/`. Read and follow its `SKILL.md` before assessment work. The skill calls this repository's reusable generator; do not offer old customer decks, validation decks, or customer-specific build scripts as alternative entry points.

The active implementation consists of:

- `generate_assessment.py`
- `ai_assess_runtime/`
- the current Excel/text input helpers and templates in the repository root and `assessment_inputs/`
- reusable assets in `assets/`
- automated checks in `test/`
- the repository-local skill in `.agents/skills/ai-assessment/`

Everything under `old/` is an archive. Do not import it, execute it, use it as a template, cite it as the current workflow, or return its generated artifacts unless the user explicitly asks to inspect historical material.

## Project mission

The skill and `generate_assessment.py` must accept different companies, services, industries, URLs and AI requirements, then generate a customer-ready editable PowerPoint. Final decks are PPTX-only and 16:9. Editable body text, tables, cards and diagram labels must be at least 12pt; the shared header and footer remain 10pt. Follow the stricter 14pt rule and the 12pt use-case-catalog exception defined by the skill.

## Implementation rules

- Implement reusable behavior in `generate_assessment.py` or `ai_assess_runtime/`. A manually edited PPTX or customer-specific post-processing script is not completion.
- Keep the implementation generic. Do not branch on a customer, company, product, service or industry name.
- Let the LLM and Web research interpret the current input. Do not embed fixed sources, recommendations, use cases or quantitative effects from archived examples.
- Preserve the supported Excel and text inputs, CLI behavior, OCI Responses API integration and authentication unless the user explicitly requests a breaking change.
- Treat percentages, monetary effects, productivity gains and management-plan figures as evidence-backed claims. Store source URLs and research decisions in the current run's research log. Omit unsupported numbers.
- Preserve the approved slide compositions defined by the `ai-assessment` skill. A request to improve content does not authorize a layout redesign.

## Required workflow

1. Enter through the `ai-assessment` skill and inspect the active generator, the current task's input, JSON and research log.
2. Modify the reusable generator path when implementation changes are requested. Do not automatically add, update or run automated tests; do so only when the user explicitly requests test work.
3. For a final real-input run, use the input supplied for the current task. Do not generate a deck for an unrelated named company solely as a generic validation artifact.
4. Run Web/LLM generation once after related content or research changes are stable. Validate the resulting JSON before rendering it.
5. Generate a PPTX when the user requests an artifact or when PPTX rendering/layout changed. Inspect all affected slides and, for layout-wide changes, the complete deck. Confirm 16:9, font floors, editable content and absence of overflow.
6. Keep customer deliverables in `output/final/`. JSON and research logs may use `output/json/` and `output/research/`. Temporary and comparison artifacts belong outside the active output tree or under `old/`.
7. Report the current input, reviewed JSON, research log, reproduction command, verification result and final PPTX. Do not present archived validation decks as deliverables.

## Verification scope

- Do not run automated test suites (including focused tests, regression tests and the full suite) unless the user explicitly requests them. This applies to routine revisions, shared implementation changes and documentation changes. Do not ask for test approval on every revision; simply skip them and report that they were not run if relevant.
- This project-specific user preference also applies when a skill recommends automated tests or test fixtures. Retain existing test files for use when explicitly requested.
- Keep verification limited to what is necessary for the requested deliverable: inspect the changed content and affected slides, and retain the generator's normal input/JSON validation. Do not expand this into unrelated regression checks.
- Do not repeat an expensive successful Web/LLM run unless a later change affects its result.
- Documentation and archive-only changes require only reading the changed instructions and checking referenced paths as needed; they do not require automated tests or new customer research.

## Completion criteria

The requested behavior must be reproducible through the repository-local `ai-assessment` skill and the active generator. The active source tree must not depend on `old/`. A one-off script, archived customer deck, hard-coded customer data or manual slide swap is not a completed implementation.
