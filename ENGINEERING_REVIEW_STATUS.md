# AZNS report review — interim safety revision

**Status: NOT a completed engineering or visual overhaul.** The existing 30 reports retain their legacy analysis and charts, which have known defects. They are not suitable as a stand-alone basis for operational or financial decisions. A warning has been placed in each document, not a certification of the remaining content.

## Verified source corrections

* `CBS AZNS-122WO1.xlsx`, `Cost Breakdown-122`, G190: plan total **USD 4,587,374.50**; H190: extra-work total **USD 1,304,440.016**. Their arithmetic sum is USD 5,891,814.516. The relationship between the two scopes, actual costs, contractual approvals and any material sheet must be reconciled before treating the sum as a like-for-like approved budget or recalculating portfolio variances. The previous unqualified claim that no budget was published is incorrect.
* `CBS AZNS-85.xlsx`, `AZNS-085 WO. Extra Work`, K54: **48.0729167** days and O54: **51.8895833** days including contingency, as labelled in that worksheet. Rows 41–49 identify rig-less steps with zero duration in that rig schedule. This does not establish zero elapsed duration or zero cost for those activities, nor does it by itself establish the complete monetary budget.
* The repeated assertion in the field-master well narratives that an operation was in a normal band on both axes has been replaced with a statement requiring separate, baseline-dependent assessment. For example, the report's own 127 WO1 values (136.6 vs 55.0 days) imply approximately 248% of planned duration; 207 (153.5 vs 80.9) implies approximately 190%. These ratios use the legacy report's values and still require reconciliation with underlying plans.

## Changes made in this interim revision

* Added a prominent review warning to all 30 documents and the source correction above to both field-master, both executive, and both language editions of the two affected well reports.
* Removed the unsupported repeated normal-band claim in the two field-master reports. The new source correction supersedes conflicting legacy statements elsewhere; those statements and legacy totals have **not** all been rebuilt.
* Centered Heading 1–3 paragraphs, running headers and footers, and first-row table paragraphs; set Times New Roman for Latin text and B Nazanin for Persian complex-script text; added missing Persian table bidi markers.
* The repair script `tools/report_safety_revision.py` documents this intervention. It is an interim patch, not the lost analytical generation pipeline, and is intended to run on the pre-revision documents, not repeatedly on already revised files.

## Outstanding — required for completion

The original request remains substantially open: full, traceable extraction of every pack, CBS and program PDF; independent reconciliation of budget scope and planned schedules; daily operational translation and formation/depth attribution; quantitative bit/BHA/hydraulics/fluids and rig-less analysis; reconstruction rather than interpolation of cost timing; replacement and visual inspection of every chart; substantive per-well and field narrative expansion; engineering validation of any recommended operating envelopes; full translation and table-width/pagination QA; and rendered-page review of all 30 reports. The existing illustrative linear cumulative cost curves and other defective charts remain embedded and must not be interpreted as observed time series. A mathematically stepped allocation based only on inferred timing would also need to be labelled as an assumption, not presented as an actual cost record. Formation assignments and numerical operating envelopes must not be invented when the source evidence is absent.

## Verification performed

The revised files were checked for ZIP integrity, XML well-formedness, embedded image presence, direct Heading 1–3 center alignment and top-level Persian table bidi markers. This is structural checking, not a visual review or validation of numerical findings. No page-by-page renderer was available in this environment. All original XLSX and PDF source files remain untouched.
