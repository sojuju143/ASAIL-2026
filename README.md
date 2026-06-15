# Measuring the Complexity of Judicial Writing
### Replication Materials — ASAIL / ICAIL 2026

Yang, K., Lim, B.L., & Madden, N. (2026). *Measuring the Complexity of Judicial Writing: An Empirical Study of Common Law Apex Courts.* Proceedings of the International Conference on Artificial Intelligence and Law (ASAIL Track).

---

## Overview

This repository contains the analysis scripts and aggregate derived metrics supporting the above paper. We study judicial writing complexity across five common law apex courts:

| Court | Jurisdiction | Period | N (substantive) |
|-------|-------------|--------|-----------------|
| High Court of Australia (HCA) | Australia | 2000–2025 | 1,483 |
| UK House of Lords (UKHL) | United Kingdom | 1996–2009 | 809 |
| UK Supreme Court (UKSC) | United Kingdom | 2009–2025 | 922 |
| Singapore Court of Appeal (SGCA) | Singapore | 2005–2025 | 1,472 |
| Singapore High Court (SGHC) | Singapore | 2005–2025 | 6,932 |
| **Total** | | | **11,618** |

**Judicial texts are not redistributed.** All shared content comprises derived numerical metrics only. Original decisions are available from the respective court websites.

---

## Repository Contents

```
scripts/
  analyze_cases_ASAIL_Public.py   — Readability and citation analysis
  pdf_parser_ASAIL_Public.py      — PDF → corpus format conversion (local PDFs only)

data/
  Judicial_Writing_Analysis_2025.xlsx  — Per-case aggregate metrics

requirements.txt
README.md
```

---

## Methodology

### Corpus format

Each cleaned judicial decision is stored as a UTF-8 plain-text file with three delimiter-separated sections:

```
-----
HEADNOTES
-----
[headnotes / catchwords]
-----
CORE JUDGMENT
-----
[full judgment text]
-----
FOOTNOTES
-----
[footnotes]
```

### Readability metrics

Metrics are computed on **CORE JUDGMENT + FOOTNOTES** combined. HEADNOTES are excluded because they do not represent judicial writing style.

#### Sentence tokenisation

NLTK Punkt is used with a custom abbreviation list (~70 entries) covering legal abbreviations — `v.`, `para.`, `s.`, `J.`, `CJ.`, `Ltd.`, `Pty.`, etc. — to suppress false sentence boundaries.

#### Citation-period neutraliser

Before tokenisation, terminal periods in citation strings are replaced with a space. This prevents citations such as `[2015] 1 AC 523.` from registering as sentence boundaries. Citations are **retained** in the text for word and syllable counting; only the orphan terminal period is neutralised.

Five patterns are applied (see `prepare_for_punkt()` in the script):

| Pattern type | Example |
|---|---|
| Bracket-year citation | `[2015] 1 AC 523.` → `[2015] 1 AC 523 ` |
| Paren-year citation | `(2004) 220 CLR 1.` → `(2004) 220 CLR 1 ` |
| Volume-reporter citation | `224 CLR 1.` → `224 CLR 1 ` |
| Paragraph pinpoint | `[45].` → `[45] ` |
| Statutory reference | `s 47B.` → `s 47B ` |

NOT neutralised: genuine sentence-ending pinpoints such as `at [45].` or `at 471.`

#### Syllable counting

The **CMU Pronouncing Dictionary** is used as the primary syllable counter. Words not found in the dictionary fall back to a vowel-group heuristic (`max(1, count of [aeiou]+ runs)`).

#### FKGL formula

```
FKGL  = 0.39 × ASL + 11.8 × ASW − 15.59
FKRE  = 206.835 − 1.015 × ASL − 84.6 × ASW
SMOG  = 3 + √(polysyllables × 30 / sentences)

where:
  ASL = words / sentences  (sentence count from neutralised text)
  ASW = syllables / words  (syllable count from original text)
```

### Citation counting

Citations are extracted and classified by jurisdiction using regex patterns for:
- Neutral citations (`[2015] UKSC 1`, `[2014] HCA 5`)
- Square-bracket law reports (`[2015] 1 AC 523`)
- Round-bracket law reports (`(2004) 220 CLR 1`)

Each citation is counted once (deduplicated by character span).

---

## Usage

### Installation

```bash
pip install -r requirements.txt

# Download NLTK resources
python -c "import nltk; nltk.download('punkt'); nltk.download('punkt_tab'); nltk.download('cmudict')"
```

### Running the analysis

```bash
python scripts/analyze_cases_ASAIL_Public.py \
  --input  /path/to/corpus/folder \
  --output results.xlsx \
  --court  HCA
```

The script walks the input folder recursively, processes every `.txt` file, and writes an Excel workbook with one row per case.

### Output columns

| Column | Description |
|--------|-------------|
| `Case` | Case name |
| `Citation` | Neutral citation |
| `Year` | Decision year |
| `Court` | Court identifier |
| `Headnotes_Words` | Word count — headnotes section |
| `Core_Words` | Word count — core judgment |
| `Footnote_Words` | Word count — footnotes |
| `FKGL` | Flesch-Kincaid Grade Level (pipeline) |
| `FKRE` | Flesch Reading Ease |
| `SMOG` | SMOG readability index |
| `Avg_Sent_Length` | Average sentence length (words) |
| `Cites_UK` … `Cites_OTHER` | Citation counts by jurisdiction |

### Converting a PDF to corpus format

```bash
python scripts/pdf_parser_ASAIL_Public.py decision.pdf output.txt
```

This reads a **local** PDF file and writes a plain-text file in the corpus format above. The headnotes section is left as a placeholder and should be filled in manually. Superscript footnote reference numbers are stripped automatically.

---

## Data

`data/Judicial_Writing_Analysis_2025.xlsx` contains the derived metrics (FKGL, FKRE, SMOG, word counts, citation counts) for the 11,618 cases in the study. Original judicial texts are not included.

---

## Citation

If you use these materials, please cite:

```
Yang, K., Lim, B.L., & Madden, N. (2026). Measuring the Complexity of Judicial Writing:
An Empirical Study of Common Law Apex Courts. Proceedings of the International Conference
on Artificial Intelligence and Law (ASAIL / ICAIL 2026).
```

---

## Licence

Scripts: MIT License
Data (derived metrics): CC BY 4.0
