# Changelog

## 27 October 2026
Edits made to the authors' full workbook before this public release.

- **Validation statuses.** The authors confirmed that all AI proposals had been checked, and the statuses were set to *Validated* for the screening decisions (sheets 05 and 06), the quality-appraisal scores (sheet 31), the taxonomy re-coding of 25 studies (sheet 28) and the transcribed performance values (sheet 32).
- **Three flags corrected** in `10_Analysis_Ready`, each with a note in `Validator_Notes`:

| Record | Flag | Was | Now | Reason |
|---|---|---|---|---|
| R0029 | Social_Media_Postprocessing | YES | NO | Only selective rejection on GenImage out-of-distribution subsets; no platform pipeline (its adversarial flag stays YES) |
| R0242 | Social_Media_Postprocessing | YES | NO | Only JPEG quality, resolution and frequency filtering; no platform pipeline |
| R0147 | Adversarial_Attack | YES | NO | Only JPEG, blur, noise and resizing; no attack on the detector |

  Dependent edits: `28_Taxonomy_Coding` robustness tags for the same three records, and `31_Quality_Appraisal` QA6 for R0242 and R0147 (1 to 0.5, benign corruptions only).
- **Effect on the reported counts.** Studies testing social-network processing 9 to 7; adversarial attack 10 to 9; platform or adversarial 17 to 15.

## Public release
Removed from the workbook for this release: the publisher-owned `Abstract` column of `04_Master_Records` (blanked), and the sheets holding working notes and intermediate synthesis. The numbers in the paper regenerate identically from this file.

## 28 October 2026
Pre-release clean-up of the public package: a template placeholder in `02_Search Development` replaced by "Not recorded"; the reviewer labels in sheets `06` and `10` now read "validated by authors" (they still said validation was required); workbook file metadata (last-modified-by) replaced with the authors' names; `supplement.pdf` is now Online Resource 1 with journal and author details; `CITATION.cff` added.
