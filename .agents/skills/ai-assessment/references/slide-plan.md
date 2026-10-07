# Preserve generator slide composition

Use the actual generator PPTX as the appearance reference. When only target assessment JSON exists, render it through the unchanged `--from-json` path into a private build directory. The resulting slide bodies, not the Oracle base's example story, govern composition.

## Page mapping

Normally keep the source page count, sequence and grouping. Optional management-plan pages depend on the target's evidence; do not borrow another company's plan or force that company's total page count.

| Source page type | Preserve in the Oracle version |
| --- | --- |
| Cover | Proposal/customer/product identity; fill the actual Oracle cover placeholders |
| Introduction | Text/image balance and explanatory blocks |
| Management plan, if present | Target company's verified figures and strategy-to-AI relationships |
| Industry/current trends | Three AI adoption developments in the end-user industry inferred from the input (not vendor product news or IR), publication timing, actual business use, service implications and visible source URLs |
| Business-impact bridge | Three independent value cards connecting the current decision issue, AI-enabled change, measurable KPI and an effect formula without unsupported effect figures |
| Use-case catalog | All 15 cases on one editable table, including selected-theme correspondence |
| PoC selection | The generator's selected themes and side-by-side reasons |
| Each PoC detail | One-page grouping of inputs/OCI roles, processing flow, business integration, evaluation and controls when the source uses one page |
| Architecture | Diagram placement, nearby theme explanations and bottom common-platform explanation |
| Cost | Total emphasis, cost table and assumptions on the same page |
| Support | Setup, responsibilities and schedule grouped as in the source |
| Closing | Closing function using the actual Oracle closing when present in the source |

Keep the opening sequence stable: cover, AI assessment introduction, verified management-plan analysis when available, industry/current-trend analysis, business-impact bridge, then the 15-use-case catalog. The industry/current-trend and business-impact pages are always present. Do not insert a separate assessment-overview page. When no verified management plan exists, the industry and business-impact pages follow the introduction directly.

Do not automatically use a fixed 12-slide outline, split each PoC into two slides, give each financial effect its own slide, or add separate governance/schedule pages. Full JSON is an evidence/content source; it does not require every nested field to appear as a new slide. Preserve the substantive coverage of the source PPTX and retain supporting detail in the source/notes.

## Layer responsibilities

- Generator reference: body arrangement, relative positions, column proportions, tables, diagrams, emphasis and density.
- This skill's `assets/base.pptx`: masters, layouts, background, title treatment, logo, footer and typography.
- Local corrections: text wrapping, spacing, box dimensions, connector clearance and faithful shortening to remove defects. Record material text changes. A repair should leave the original composition recognizable.

Use the base's actual content layout and retain its brand furniture. Derive a clean body canvas by removing identified example objects, then transfer/reproduce the generator's body composition. Reusing source coordinates is permitted and preferred when they preserve the requested appearance. Correct inherited clipping rather than reproducing it.

## Restyling helper

`python scripts/run_assessment.py restyle SOURCE_PPTX OUTPUT_PPTX --plan PLAN_JSON` uses native OOXML and preserves imported media/relationships. It requires lxml. Keep the output separate from its sources.

The plan has `slides`, in source order. Each entry includes `source_slide`, `kind` (`cover`, `body`, `closing`), `title` and optional `title_size`. The three-page base maps these kinds to pages 1, 2 and 3, including when an older plan has stale `base_slide` numbers. Cover entries include `subtitle`. Presenter names are omitted by default; use `show_presenter: true` and `presenter` only when explicitly requested. Optional `remove_shapes` lists inspected source IDs for furniture; `shape_edits` maps source IDs to `text`, `size`, `geometry` (x/y/width/height in inches) or `line_spacing_pt`. Recheck IDs against the exact input deck.

This helper targets the inspected generator's 16:9 furniture conventions and current Oracle content placeholders. It is not a universal arbitrary-deck converter. It retains source body objects and uses a separate `.layout-map.json` for correspondence. Source raster illustrations remain images; verify and disclose diagram-label editability. Native visual inspection and local plan corrections are required after conversion.

## Fidelity review

Compare pages side-by-side: page count/grouping, object placement, text coverage, chart/table structure and priority IDs. Confirm that only intended brand changes and necessary local repairs alter the source appearance. Check that unagreed planning figures remain labeled as assumptions and that costs retain their scope and price limitations. Show OCI service costs and totals as comma-separated integer yen amounts; do not round them into approximate ten-thousand-yen labels.

## Assessment typography and photo alignment

Minimum editable body, card, table and diagram text is 14pt. The single 15-use-case table may stay 12pt. Both industry/current-trend and management-plan source URLs use the same gray 9pt bottom line above the footer. The cover and introduction title is `AI Use Case Assessment`, without an OCI prefix. Add `<formal company name>様のAI構築をご支援` as a dark 16pt introduction subtitle, retaining the legal company name and a gap above the explanatory blocks. The management-plan bottom band summarizes input-derived AI application direction, rather than caveats. Keep shared template header/footer settings. Allocate more height and spacing and faithfully shorten repeated wording before considering page regrouping. Do not use automatic font shrinking to bypass the minimum.

For a right-half introduction photo with a band, align both to the same half-slide edges. Use proportional fill/crop; do not inset the photo while leaving the band full width.

For an existing output, `scripts/revise_pptx.py SOURCE EDIT_PLAN OUTPUT` changes only affected slide XML. The plan supports minimum_font_pt, preserve_pages, and per-page shapes keyed by the inspected shape ID. Shape edits support geometry, text, font_pt, line_pt, wrap, remove, table cells, row_heights, column_widths, vertical_anchor and alignment. An optional document_title updates the PPTX core title. Keep this edit plan as the reproducible follow-up to the initial restyling outline.

The three technical detail pages lead directly to architecture; omit the standalone existing-system/AI/OCI role-sharing slide. Center the selection numbers in their circles. Align the cost total label and amount vertically, and anchor the three assumption notes below the measured table bottom with visible spacing.
