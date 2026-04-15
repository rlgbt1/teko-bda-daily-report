# BDA Test Sample Audit Report - Current Stage and Fix Plan

**Audited output:** `output/test_sample_01042026.pptx`  
**Reference files:** `RESUMO DIARIO-01-12-2025 2.pdf`, `test_sample_01042026.pdf`  
**Template:** `assets/template_v1.pptx`  
**Last reviewed:** 15 April 2026  
**Status:** template copy is implemented. Slide 3 preserves template groups. The visible overlap/extra-note blockers on slides 4, 5, and 11 have been corrected; slides 4-11 still need broader template-object correction.

---

## Executive Verdict

This is not hopeless, but it is still in a rough visual QA stage. The biggest architectural problem has improved: `pptx_builder.py` now starts from `assets/template_v1.pptx` with `shutil.copy(...)`. Slide 3 now keeps the template's grouped KPI objects and updates their text in-place. However, the generator still clears slides 4-11 and recreates those slides with manual rectangles, text boxes, ovals, and chart pictures.

That explains the current symptoms:

- tables do not keep the exact original/master-template sizes because many native OLE/table shapes are still deleted;
- slide 4 ovals previously sat on top of the LUIBOR area; they have been moved to the right-side free zone;
- slide 5 previously had generated chart images over the P&L table; P&L is now constrained to the lower-left template-sized zone;
- slide 9 loses the original table column proportions because the rebuilt table uses equal-width columns;
- slide 11 previously had an extra bottom note; the generated `Nota:` band has been removed.

The project is therefore past the "missing template" stage, but not yet at "template-faithful generation." The next pass should be layout correction plus selective template-preservation.

---

## Current Code Stage

| Area | Stage | Evidence | Meaning |
|---|---|---|---|
| Template loading | Improved | `BDAReportGenerator.build()` copies `assets/template_v1.pptx` before saving | Cover, agenda, slide size, and base theme are now available |
| Slide 1 | Mostly safe | only date shape is updated | good pattern |
| Slide 2 | Safe | skipped/static | good pattern |
| Slide 3 | Improved | template groups are preserved and updated in-place | icons, connectors, central oval, and group coordinates now match the template structure |
| Slides 4-11 | Still risky | `_clear_slide(slide)` runs before each rebuild method | native table sizes, OLE chart placeholders, and original coordinates are discarded |
| Slides 4-5 | Improved | slide 4 ovals moved off the table zone; slide 5 P&L constrained to lower-left | overlap blockers fixed; full visual QA still needed |
| Slides 8-9 | Not template-native | output uses rectangle/textbox grids instead of template tables | column widths and table proportions drift |
| Slide 11 | Improved | bottom `Nota:` band removed | no extra footnote-like shape remains in generated PPTX |

---

## Root Cause Update

The old root cause was: "the deck was built from scratch instead of loading the template."

The current root cause is more specific:

**The deck now loads the template and preserves slide 3, but the generator still clears slides 4-11 and rebuilds them as manual Pattern B shapes. This preserves the file-level template but not the slide-level layout objects that define the original sizes and positions on the remaining problem slides.**

Required architectural correction:

1. Keep loading the template with `shutil.copy("assets/template_v1.pptx", output_path)`.
2. Stop clearing every slide by default.
3. For each slide, choose one of three modes:
   - **Preserve static slide:** slide 1 date-only, slide 2 untouched.
   - **Patch native template objects:** slides with usable tables/groups should keep those objects and only update text/cell contents.
   - **Controlled rebuild:** only delete OLE/chart objects that cannot be edited, then replace them with generated chart images at the template placeholder coordinates.

---

## Priority Bugs From Latest Visual QA

### Slide 4 - Liquidez MN

**Severity:** critical  
**User-visible issue:** ovals were on top of tables.

The template places the KPI ovals around:

- `Liquidez Total`: `L=2.68 T=5.38 W=1.10 H=0.86`
- `Juros Diario`: `L=4.20 T=5.40 W=1.10 H=0.86`

Those coordinates are valid only when the original OLE/table blocks keep their template layout. In the rebuilt output, the LUIBOR table flowed into the same vertical band, so the ovals collided with rows near the bottom of that table.

**Fix direction:**

- Done: moved the two ovals to the right-side free zone. Geometry check shows `0.0` overlap with the left table zone.
- Better: preserve the template's main table blocks/placeholder geometry and only replace OLE charts with images.
- Also restore slide 4 chart placeholders:
  - LIQUIDEZ BDA pie/chart area at about `L=9.26 T=0.44 W=4.01 H=3.25`
  - LUIBOR trend chart at about `L=9.00 T=5.24 W=4.33 H=1.24`

### Slide 5 - Liquidez MN Cash Flow

**Severity:** critical  
**User-visible issue:** two pies/charts were sitting on top of the P&L table.

The code inserts generated chart images at:

- desembolsos chart: `L=6.21 T=3.34 W=2.69 H=2.53`
- reembolsos chart: `L=8.52 T=3.45 W=4.09 H=2.40`

The template expects the P&L/OLE table around the lower-left zone, roughly `L=0.40 T=5.60 W=5.77 H=0.82`. The rebuilt P&L table was full-width and started immediately after the cash-flow table, so it collided with both charts.

**Fix direction:**

- Done: P&L Control is now constrained to the lower-left zone and uses compact row heights.
- Done: geometry check shows `0.0` overlap between P&L text/header shapes and chart pictures.
- Still needed: visual QA against exported PDF/image.

### Slide 9 - Operacoes BDA

**Severity:** major  
**User-visible issue:** table formatting is not following the original sizes.

The template has a native `Table 161` at about `L=1.75 T=3.87 W=9.00 H=2.26`. The generated output rebuilds the carteira table with eight equal columns from `L=0.30` across almost the full slide width. That makes narrow columns too wide and monetary columns too cramped or visually wrong compared with the master.

**Fix direction:**

- Preferred: fill template `Table 161` instead of rebuilding the carteira grid.
- If rebuilding remains necessary, use explicit non-equal widths:
  - `Cód.` narrow
  - quantities medium
  - `Montante`, `J. Anual`, `J. Diario` wider
- Add/restore the parent section label `Carteira De Titulos` above the sub-sections.
- Keep KPI ovals in the bottom band and away from data rows.

### Slide 11 - Informacao de Mercados 2/2

**Severity:** major  
**User-visible issue:** one extra footnote below the tables.

This was generated by code. `_slide_market_info_2()` added a full-width bottom note band:

- rectangle at `L=0.25 T=5.60 W=12.83 H=0.85`
- text box with `Nota: ...` at `L=0.35 T=5.66`

The template/reference already uses right-side commentary/note areas. The extra full-width bottom band created a duplicate footnote-like block under the tables.

**Fix direction:**

- Done: removed the bottom `Nota:` band from slide 11.
- Done: regenerated PPTX contains no `Nota:`/footnote-like bottom shape.

---

## Slide-by-Slide Status

| Slide | Current status | Main issue | Fix priority |
|---|---|---|---|
| 1 Cover | acceptable | date-only update works | low |
| 2 Agenda | acceptable | static template slide | low |
| 3 Sumario Executivo | improved | template groups/icons/connectors preserved; visual QA still needed for text styling | low/medium |
| 4 Liquidez MN 1/2 | improved | ovals no longer overlap table; missing/right-side charts still need template-faithful placement | major |
| 5 Liquidez MN 2/2 | improved | chart images no longer overlap P&L; final visual QA still needed | major |
| 6 Liquidez ME | mostly workable | chart coverage still needs final visual QA | medium |
| 7 Mercado Cambial | workable | chart depth/series still simplified | medium |
| 8 Mercado Capitais | partially workable | should preserve/fill native table or match its dimensions | medium |
| 9 Operacoes BDA | needs correction | equal-width rebuilt table ignores template `Table 161` | major |
| 10 Markets 1/2 | workable | missing monthly history columns if required by final spec | medium |
| 11 Markets 2/2 | improved | extra generated `Nota:` band removed | medium |

---

## Corrected Fix Plan

### Phase 1 - Fast Layout Stabilization

Estimated time: **2-3 hours**

1. Done: Slide 4 KPI ovals moved out of the LUIBOR table collision zone.
2. Done: Slide 5 P&L Control restricted to the lower-left template area.
3. Done: Slide 11 generated bottom `Nota:` band removed.
4. Regenerate PPTX and visually inspect slides 4, 5, 9, and 11.

This should produce a much better version quickly, even before the full template-native rewrite.

### Phase 2 - Template-Faithful Tables

Estimated time: **4-6 hours**

1. Slide 9: use template `Table 161` if practical, otherwise rebuild with measured template-like column widths.
2. Slide 8: align table dimensions to the template table/object sizes.
3. Slide 4: decide whether to preserve original OLE table zones or continue with a controlled rebuild.
4. Add a small layout audit script that fails if shapes overlap protected table zones.

### Phase 3 - Visual QA and Polish

Estimated time: **2-4 hours**

1. Export PPTX to PDF.
2. Compare slides 4, 5, 9, and 11 against the reference PDF.
3. Tune font sizes, row heights, label positions, and chart bounding boxes.
4. Confirm there are no shapes below table sections except the standard footer.

---

## Time Estimate

For the specific problems reported here:

- **Good corrected version:** about **3-5 focused hours**.
- **Template-faithful version suitable to proceed confidently:** about **1 working day**.
- **Full hardening with overlap checks and visual QA workflow:** **1-2 working days**.

The fastest path is to fix slides 4, 5, 9, and 11 first, generate a new deck, and review only those slides before spending more time on the less visible polish.

---

## Concrete Code Notes For The Next Developer

In `src/report_generator/pptx_builder.py`:

- `build()` already copies the template. Do not remove that.
- The risky line is the loop that calls `_clear_slide(slide)` for slides 4-11.
- Slide 3 now skips `_clear_slide(slide)` and updates the template groups in-place.
- Slide 4 ovals are now created in the right-side free zone near `L=9.25` and `L=10.78`.
- Slide 5 chart placement is in `_add_pie_charts_mn()`.
- Slide 5 P&L is now constrained to roughly `L=0.407 T=5.600 W=5.770`.
- Slide 9 carteira columns are equal-width because `ct_w_each = (SLIDE_W - Inches(0.5)) / n_ct`.
- Slide 11 extra bottom note block has been removed from `_slide_market_info_2()`.

Immediate patch completed:

1. Done: Slide 5 P&L width is approximately `5.77"` and starts around `T=5.60`.
2. Done: Slide 5 chart images remain right of `L=6.20`.
3. Done: Slide 11 bottom note block removed.
4. Still pending: replace Slide 9 equal column widths with measured widths or template table filling.
5. Still pending: export the regenerated PPTX to PDF and inspect visually.
