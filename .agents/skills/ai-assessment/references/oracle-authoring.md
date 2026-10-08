# Oracle PowerPoint authoring for AI assessment

This file contains the Oracle presentation rules needed by `ai-assessment`.
The separate `oracle-ppt-creator` skill is not a dependency of this workflow.

## Brand source and page roles

- Use `../assets/base.pptx` from this skill. It is a 16:9 lightweight copy of the retained Oracle design, with the native cover, body furniture and closing.
- Page 1 is the cover, page 2 is the body, and page 3 is the closing. Resolve the actual presentation order through `ppt/presentation.xml` and its relationships when inspecting or rebuilding the asset.
- Preserve the template's real master, layouts, background, Oracle logo, title treatment, slide-number field, copyright footer and cover affiliation. Do not recreate brand furniture as independent text boxes or flatten slides into pictures.
- The generator PPTX determines body composition, section order, tables, diagrams and information density. The base supplies brand furniture and typography. Use the closing only when the generator's source deck includes a closing.
- Keep the original base as an asset. For a customer deck, render to a new file; do not replace the skill template with an output deck.

## Content and formatting

- Default to the warm palette in `ai_assess_runtime/presentation_theme.py`: charcoal text, Oracle red / terracotta / muted gold accents, pale warm panels and warm table headers. Preserve these generated body colors during restyling. Prefer square cards and fine rules; retain the actual Oracle master and raster artwork.

- Write direct Japanese business prose with enough explanation to stand on its own. Use varied comparisons, processes, diagrams and tables as warranted by the generator; do not force repeated cards or slogans.
- Keep equal conditions at equal visual weight. Use Oracle red for meaningful emphasis. Do not invent customer achievements, sources, percentages or costs. Keep facts, verified external evidence and proposals distinguishable.
- Use Meiryo UI for newly authored Japanese, Latin and numeric text, including East Asian font properties. The normal cover title is about 30pt, cover subtitle about 18pt and body title about 24pt. The assessment's stricter rule governs editable body, table, card and diagram text: at least 14pt, except the 15-case catalog at 12pt and the specified source URLs at 9pt. Preserve inherited footer settings, normally 10pt.
- Fit text by shortening faithfully, adjusting boxes, spacing and wrapping, then inspect the result. Do not shrink automatically below the applicable floor. Reserve the inherited footer area.
- Preserve image proportions. For the introduction's right-half photo and overlay band, align their left and right edges and use proportional fill/crop.
- Omit a presenter from the cover unless the current request explicitly asks for one. Keep the legal company name, service name and affiliation. This AI assessment rule overrides a generic Oracle deck's presenter default.

## Restyling and focused revision

- Use `scripts/restyle_pptx.py` from this skill with the reviewed generator reference PPTX, this skill's `assets/base.pptx`, and a reviewed page/shape plan. Keep the output separate from both inputs. The helper preserves source body objects and their media while replacing template sample body content.
- Use cover/body/closing page roles 1/2/3. The restyler also maps old plan page numbers by role when given this three-page base.
- For an existing output, inspect the exact latest file and its displayed page order. Use `scripts/revise_pptx.py` with a reviewed shape edit plan for focused changes. Preserve unrelated slide XML, media, notes, hidden flags, masters, layouts and metadata.
- Retain the base slide's title, `sldNum` and `ftr` placeholders on every body page, including all three technical detail pages. Do not repair a missing footer by adding standalone text.
- Keep source raster illustrations as images and describe their editability honestly. Native text, tables, cards and labels should remain editable.

## Verification

- Validate a staged candidate before delivering it. Check ZIP/XML integrity, every internal relationship target, 16:9 dimensions, slide order and count, hidden flags, metadata title, source coverage, editability and font floors.
- Resolve slide order via `ppt/presentation.xml`; a `slideN.xml` filename need not be displayed page N. A slide with no local shapes can still display its closing design through a master or layout.
- For a focused edit, compare changed slide parts with the source and confirm unrelated package parts remain unchanged. Different ZIP compression sizes alone are not content changes.
- Prefer native Microsoft PowerPoint rendering for visual review when available. Export the staged presentation, identify that exact presentation before export, and do not close or overwrite unrelated open documents. Temporary PDF/PNG files are fine for QA; the customer deliverable remains PPTX.
- Compare rendered page count with visible-slide count because hidden pages may be omitted from PDF export. Inspect every page of a new deck at readable size for text wrapping, clipping, alignment, title/subtitle separation, connector clearance, logo and footer. A contact sheet helps but does not replace close inspection of affected pages.
- Confirm the three technical detail pages and adjacent pages have matching inherited footers. Confirm the PDF title matches the customer-facing PPTX title when exporting PDF for QA.
- Report only checks actually performed. If native rendering is unavailable, state that limitation. After validation, copy the candidate to `output/final/` and verify it matches the reviewed candidate.
