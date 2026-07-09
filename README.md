# Australian Curriculum v9 (F–6) content descriptions — clean JSON

All **614 content descriptions** of the Australian Curriculum version 9 for
Foundation to Year 6, across all seven F–6 subjects, as one flat,
developer-friendly JSON file.

```json
{
  "code": "AC9MFN01",
  "learningArea": "Mathematics",
  "subject": "Mathematics",
  "strand": "Number",
  "year": "F",
  "description": "name, represent and order numbers including zero to at least 20, using physical and virtual materials and numerals",
  "years": ["F"]
}
```

## Why this exists

ACARA publishes the Australian Curriculum as DOCX downloads and — since
2024 — as the official **Machine Readable Australian Curriculum (MRAC)**
in RDF/XML, JSON-LD and SPARQL form
(<https://www.australiancurriculum.edu.au/machine-readable-australian-curriculum>).
MRAC is the authoritative machine-readable release, and if you need the
full semantic model you should use it.

What has been missing is the simple thing: a **flat JSON file** you can load
in five minutes to answer questions like *"give me every Year 3 Mathematics
content description with its code."* This repository provides that view,
parsed directly from ACARA's official DOCX downloads, with a
reproducibility script so you can regenerate and verify it yourself.

## Contents

| Subject | Content descriptions |
|---|---|
| English | 189 |
| Mathematics | 139 |
| HASS F–6 (Humanities and Social Sciences) | 84 |
| Science | 77 |
| Health and Physical Education | 59 |
| Digital Technologies | 38 |
| Design and Technologies | 28 |
| **Total (F–6)** | **614** |

## Record schema

| Field | Type | Meaning |
|---|---|---|
| `code` | string | Official AC9 content-description code, e.g. `AC9M3N02` |
| `learningArea` | string | Learning area (`Technologies` covers both Technologies subjects; `Humanities and Social Sciences` covers HASS F–6) |
| `subject` | string | Subject as published (`English`, `Mathematics`, `Science`, `HASS F-6`, `Health and Physical Education`, `Digital Technologies`, `Design and Technologies`) |
| `strand` | string | Strand the description sits under, e.g. `Number`, `Literacy`, `Knowledge and understanding` |
| `year` | string \| null | Single year level (`"F"`–`"6"`) when the description belongs to one year; `null` for banded subjects (HPE, Digital Technologies, Design and Technologies use year bands) |
| `description` | string | The content description text |
| `years` | string[] | All year levels covered, e.g. `["F"]`, `["3"]`, `["1","2"]` |

## Data notes

- **Year levels come from document placement**, not from decoding the AC9
  code string: each description carries the year label of the table ACARA
  placed it under ("Foundation", "Year 3", "Years 1–2"). This matters for
  edge cases — e.g. `AC9TDIFP01` sits under Foundation, which a
  code-pattern heuristic can misread as an F–2 band.
- **Equation fractions are rendered as text** (`1/2`, `1/10`): ACARA embeds
  some fractions as Word equation objects, which naive extraction drops
  (see `AC9M3N02`).
- Unicode whitespace variants (non-breaking, thin and narrow no-break
  spaces) in the source documents are normalised to plain spaces.
- Content *elaborations*, achievement standards, general capabilities and
  cross-curriculum priorities are **not** included — this dataset is
  content descriptions only, Foundation to Year 6 only.

## Reproducing the dataset

1. Download the seven official DOCX files (URLs are documented in
   `DOCX_SOURCES` inside `scripts/parse_acara_docx.py`) into a directory,
   named `english.docx`, `mathematics.docx`, `science.docx`,
   `health-and-physical-education.docx`, `hass.docx`,
   `digital-technologies.docx`, `design-and-technologies.docx`.
2. `pip install python-docx`
3. Regenerate and verify:

```
python scripts/parse_acara_docx.py <docx-dir> -o regenerated.json
python scripts/parse_acara_docx.py <docx-dir> --verify data/acara-v9-f6-content-descriptions.json
```

The `--verify` run asserts that a fresh parse of the official documents
reproduces this repository's dataset exactly (all 614 records, all fields,
order-independent).

## Provenance

Parsed **2026-07-09** from the official curriculum-content DOCX files
published on the [Australian Curriculum v9 website](https://v9.australiancurriculum.edu.au/)
(download URLs recorded in the parser). No content has been reworded;
descriptions appear exactly as published, apart from the whitespace
normalisation and fraction rendering noted above.

## Licence and attribution

- **Data** (`data/`): the content descriptions are from the
  **Australian Curriculum version 9, © ACARA (Australian Curriculum,
  Assessment and Reporting Authority)**, licensed under
  [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)
  (full text in `LICENSE-DATA`). This repository redistributes them with
  attribution under the same licence.
- **Scripts** (`scripts/`): MIT licence (see `LICENSE`).

**Disclaimer:** this is an unofficial extraction and is not endorsed by
ACARA. For compliance-critical uses (school programming, assessment,
reporting), verify against the official curriculum at
[australiancurriculum.edu.au](https://www.australiancurriculum.edu.au/) or
the official MRAC release.

## Related

- **[Primary Learning Map (Australia)](https://learningmap.au)** — a
  browsable site built on this dataset and the open Marble Skill Taxonomy:
  one page per topic with prerequisite chains, mastery evidence, and
  candidate AC9 alignments.
- **[Marble Skill Taxonomy](https://github.com/withmarbleapp/os-taxonomy)**
  — 1,590 primary-years micro-topics in a prerequisite graph; this dataset
  has been proposed there as the `acara-v9` standards source.

## Changelog

- **1.0.0** (2026-07-09) — initial release: 614 records parsed from the
  official ACARA v9 F–6 curriculum-content DOCX downloads.
