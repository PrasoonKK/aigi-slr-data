# Replication package: AI-generated image detection, a systematic literature review

Protocol, coded workbook, scripts and supplementary tables for **"Towards Universal AI-Generated Image Detection: A Systematic Literature Review of Methods, Generalization, Robustness, and Open Challenges"** by Prasoon Kumar Kashyap and Aryan Jha (Manipal Institute of Technology Bengaluru, Manipal Academy of Higher Education).

The review follows PRISMA 2020. It covers peer-reviewed journal articles on AI-generated image detection published between January 2020 and August 2026 (Scopus, searched on 13 August 2026): 473 records identified, 382 screened, 99 included.

## What is here
| Path | Content |
|---|---|
| `protocol/AIGI-SLR-Protocol.docx` | The written review protocol (objective, research questions, eligibility criteria, search strategy, appraisal rubric). |
| `AIGI Workbook.xlsx` | The coded workbook (details below). |
| `scripts/gen_from_workbook.py` | Regenerates every statistic, table row, figure input and corpus reference in the paper from the workbook. |
| `scripts/crossref_cache.json` | Volume, issue and page data fetched from Crossref by DOI, so the script needs no network. |
| `supplement.tex`, `supplement_tables.tex`, `supplement.pdf` | Online Resource 1 (supplement to the article): study-level coding of the 99 included studies (Table S1), every transcribed performance value (S2), full-text exclusions (S3), reports not retrieved (S4), and the verbatim Scopus search strings. |
| `generated/` | The script outputs the supplement needs (tracked); the rest is rebuilt on demand. |
| `refs_corpus.bib`, `refs_background.bib` | Bibliographies for the supplement (corpus entries are generated from the workbook). |
| `CHANGELOG.md` | Edits made to the workbook on 27 October 2026. |

## The workbook
| Sheet | Content |
|---|---|
| `00_Protocol_Changes` | Dated log of protocol amendments |
| `01_Search_Log`, `02_Search Development` | The two Scopus searches (S01, S02) and the five pilot searches |
| `03_Protocol` | Protocol text |
| `04_Master_Records` | All 382 unique records after deduplication |
| `05_Title_Abstract_Screening`, `06_Full_Text_Screening` | Every screening decision with its reason and exclusion code |
| `10_Analysis_Ready` | The 72-field extraction matrix for the 99 included studies |
| `14_Generalization_Map` | Per-study map of generalization tests (used by the integrity checks) |
| `27_Final_Taxonomy_Framework`, `28_Taxonomy_Coding` | The five-axis taxonomy and the coding of each study |
| `30_Integrity_Checks` | Live formulas that check the record sets, the PRISMA arithmetic and consistency between flags and derived fields |
| `31_Quality_Appraisal` | Scores for the eight appraisal questions (QA1 to QA8) |
| `32_Reported_Results` | The 197 performance values quoted from the extraction records, each with its verbatim source fragment |

**Not included:** the full texts of the studies (publisher copyright), and the abstracts of the Scopus records. The `Abstract` column in `04_Master_Records` is intentionally empty; re-export it from Scopus with the search strings in `01_Search_Log` if needed. Working notes used while planning the review are also left out.

## Validation status and use of AI
Large-language-model tools were used to propose screening decisions and pre-fill extraction fields (Gemini Pro, Google; exact version not recorded), and to score the quality-appraisal items QA1, QA3 and QA8, transcribe reported values, write the analysis code and prepare the manuscript (Claude, Anthropic, accessed through Claude Code). Every AI proposal was recorded in the workbook with a validation status, and the authors validated all of them (statuses updated on 27 October 2026, see `CHANGELOG.md`). Decisions were not made independently by two reviewers, so no inter-rater agreement is reported.

## Reproducing the numbers
Requires Python 3 and `openpyxl`.
```bash
pip install openpyxl
python3 scripts/gen_from_workbook.py "AIGI Workbook.xlsx"
```
This rewrites `generated/` and `refs_corpus.bib`. To rebuild the supplement: `tectonic -X compile supplement.tex`, or pdfLaTeX with BibTeX.

The workbook formulas (sheets `14`, `30`, `31`) recalculate when the file is opened in Excel.

## Citation and licence
Citation: to be added once the paper is published.

Licence: the script and LaTeX files are under the **MIT License**; the data and documents (workbook, protocol, supplement, generated tables) are under **CC BY 4.0**. See `LICENSE`, `LICENSE-MIT.txt` and `LICENSE-CC-BY-4.0.txt`.
