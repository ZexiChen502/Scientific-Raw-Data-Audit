# Raw Data Audit Tool

This folder contains a lightweight raw-data audit utility for screening tabular scientific datasets before public release, review, or internal quality control.

If you use this tool in academic work, please cite or acknowledge this repository.

The public version runs from the command line. It reads heterogeneous raw-data files, applies statistical and structural screening checks, and produces a Word report plus an optional JSON summary.

## Important interpretation note

This software is a **screening and review-prioritization tool**, not a misconduct detector.

A red, yellow, or green label is not proof that a dataset is correct or incorrect. The labels only indicate how strongly the file triggered statistical integrity indicators and how much manual review may be needed. Legitimate scientific datasets can contain repeated values, monotonic trends, rounded values, ranks, time courses, dose-response series, or sorted tables. The tool therefore includes context-aware rules to reduce false positives.

## Files

```text
raw_data_check_soft/
  raw_data_audit.py          # Core CLI and audit engine
  requirements.txt           # Python package requirements
  simulation_data/           # Small demonstration datasets
```

Generated outputs such as reports, executable builds, and PyInstaller work directories are intentionally excluded from version control.

## What the tool checks

The audit combines several independent families of checks.

### 1. Readability and file-format handling

The tool checks whether each file can be discovered and read using the available Python dependencies. Files that require optional dependencies are recorded clearly in the report instead of causing the whole audit to fail.

### 2. Duplicate and copy/paste indicators

The tool checks for:

- exact duplicate rows;
- duplicate rows after whitespace and missing-value normalization;
- duplicate column names;
- contiguous repeated row blocks;
- identical byte-level file contents across files.

These checks are intended to identify possible copy/paste, file duplication, or table assembly issues. Repeated sample IDs in long-format data are not treated as duplicate rows by themselves.

### 3. Terminal-digit and rounding patterns

For measurement-like numeric columns, the tool checks for:

- excessive terminal `0` or `5` digits;
- unusually low terminal-digit entropy;
- concentrated decimal endings;
- strong rounding patterns.

These checks are inspired by statistical audit approaches in which manually rounded or fabricated measurements may show non-random terminal digits. They are only applied when there are enough numeric observations and are downgraded or skipped for columns where such patterns may be expected.

### 4. Heaping and low-randomness signals

The tool checks for:

- unusually low unique-value ratios;
- excessive repeated exact values;
- high mode frequency;
- continuous-looking columns with strong heaping.

These signals can occur naturally in count data, categorical scores, ranks, and rounded measurements, so they are treated as review indicators rather than direct evidence of data problems.

### 5. Table-coordinate regularities

The tool checks whether numeric values change too regularly with table position, including:

- arithmetic progressions down rows;
- repeated adjacent-cell deltas;
- strong correlation with row order;
- numeric blocks that can be explained by row and column indices;
- row/column gradient patterns.

This check is designed to catch suspiciously regular table structures, while avoiding over-penalizing legitimate time courses, dose-response tables, sorted rankings, or training curves.

### 6. Benford-style first-digit checks

Benford-style checks are applied only when appropriate:

- positive values;
- enough observations;
- values span multiple orders of magnitude;
- values are not bounded metrics, p-values, ranks, correlations, percentages, or coordinates.

Benford deviations alone are treated as weak evidence and normally produce yellow-level findings only.

### 7. Scientific sanity checks

Column names are used to apply basic domain-aware bounds:

- p-values, adjusted p-values, q-values, and FDR values should be in `[0, 1]`;
- correlations should be in `[-1, 1]`;
- probabilities and AUC values should be in `[0, 1]`;
- count-like columns should not contain negative values;
- infinite numeric values are flagged;
- severe missingness is summarized.

### 8. Context-aware false-positive control

The tool infers column roles from names and value patterns. The following columns are skipped, downgraded, or interpreted cautiously where appropriate:

- identifiers: sample, patient, cell, gene, drug, barcode, group labels;
- dose, concentration, time, day, epoch, step, iteration, rank;
- t-SNE, UMAP, PCA, coordinate columns;
- statistical summaries such as p-values, correlations, R2, loss, scores;
- count-like columns;
- low-cardinality columns.

This is important because many legitimate raw-data tables are structured by design.

## Supported input formats

Currently supported or partially supported formats:

| Format | Reader | Notes |
|---|---|---|
| `.csv` | `pandas.read_csv` | Fully supported |
| `.tsv` | `pandas.read_csv(sep="\t")` | Fully supported |
| `.txt` | delimiter sniffing + pandas | Supports comma, tab, semicolon, pipe when detected |
| `.xlsx` | `pandas.read_excel` + `openpyxl` | One audit table per sheet |
| `.xls` | `pandas.read_excel` | May require additional Excel engine support |
| `.Rdata`, `.RData`, `.rda`, `.rds` | `pyreadr` | One audit table per readable R object |
| `.parquet` | `pandas.read_parquet` | Requires parquet backend such as pyarrow or fastparquet |
| `.h5ad` | `anndata` | Audits `obs` and `var`; avoids densifying large matrices |
| `.h5`, `.hdf5` | `h5py` | Metadata-first; audits small numeric 1D/2D datasets |

Missing optional readers are recorded in the report as skipped or metadata-only files.

## Installation

Recommended full installation:

```bash
pip install -r requirements.txt
```

Equivalent manual installation:

```bash
pip install pandas numpy scipy python-docx openpyxl pyreadr anndata h5py
```

Minimum installation for CSV/TSV/TXT auditing and Word reports:

```bash
pip install pandas numpy scipy python-docx
```

Optional format support:

```bash
pip install openpyxl   # Excel .xlsx
pip install pyreadr    # RData/RDS/RDA
pip install anndata    # H5AD
pip install h5py       # H5/HDF5
```

## Command-line usage

From the repository root:

```bash
python raw_data_check_soft/raw_data_audit.py --raw-dir raw_data
```

Specify output paths:

```bash
python raw_data_check_soft/raw_data_audit.py \
  --raw-dir raw_data \
  --output-docx raw_data_check_soft/raw_data_audit_report.docx \
  --output-json raw_data_check_soft/raw_data_audit_report.json
```

Run only the built-in demonstration datasets:

```bash
python raw_data_check_soft/raw_data_audit.py \
  --run-simulations \
  --output-docx raw_data_check_soft/raw_data_audit_simulation_report.docx \
  --output-json raw_data_check_soft/raw_data_audit_simulation_report.json
```

Run simulations first and then audit real raw data:

```bash
python raw_data_check_soft/raw_data_audit.py \
  --run-simulations \
  --audit-simulations-and-raw \
  --raw-dir raw_data
```

Audit only selected formats:

```bash
python raw_data_check_soft/raw_data_audit.py \
  --raw-dir raw_data \
  --formats csv,rdata,xlsx
```

## Graphical interface

A separate graphical interface can be built on top of the same audit engine, but the public GitHub version in this folder only documents and distributes the command-line tool.

The core function `run_audit()` and progress callback support in `raw_data_audit.py` can be reused by private GUI wrappers or downstream applications.

## Demonstration data

The folder `simulation_data/` contains small synthetic datasets used to demonstrate expected behavior and verify the audit logic.

The simulation mode creates or refreshes these files:

| File | Purpose | Expected behavior under conservative settings |
|---|---|---|
| `sim_clean_continuous.csv` | Random continuous measurements | Green |
| `sim_benign_time_dose.csv` | Legitimate monotonic/time/dose-style data | Green or low yellow only |
| `sim_duplicate_blocks.csv` | Large duplicated row blocks | Yellow under conservative red rules |
| `sim_terminal_digit_heaping.csv` | Excess `.0`/`.5` or terminal digit concentration | Yellow/red depending thresholds |
| `sim_coordinate_gradient.csv` | Numeric block tied to row/column indices | Yellow under conservative red rules |
| `sim_low_randomness_heaped.csv` | Continuous-looking values drawn from a small repeated set | Yellow |
| `sim_multisheet.xlsx` | Multi-sheet Excel example | Yellow when Excel support is available |

These datasets are intentionally small and are useful for demonstrating the report format, progress callbacks, and scoring behavior.

## Output files

The tool can generate:

- a Word report (`.docx`), intended for human review;
- a JSON report (`.json`), intended for programmatic inspection or debugging.

The Word report includes:

1. audit function and principles;
2. executive summary;
3. dependency availability;
4. simulation validation results if simulations are run;
5. top findings;
6. red/yellow/green file sections;
7. skipped or metadata-only files;
8. methods appendix.

## Red/yellow/green scoring rules

Each finding has:

- a check family;
- a severity label;
- a confidence label;
- a numeric score;
- a message;
- optional false-positive notes.

The file-level score is grouped by check family to avoid over-counting many related findings from the same issue.

Current check families include:

```text
readability
duplicates
digit_patterns
heaping_low_randomness
coordinate_regularities
benford
scientific_bounds
cross_file_similarity
missingness_structure
```

Within each family:

1. the strongest finding contributes its full score;
2. up to two additional findings in the same family contribute only 25% of their score.

This prevents a single problematic column from creating many near-duplicate penalties.

### Default category thresholds

The default conservative thresholds are:

```text
Green:  score < 25
Yellow: score >= 25
Red:    stricter priority-review category
```

A file is marked **red** only if one of the following conditions is met:

```text
score >= 90
OR red_count >= 3 and score >= 70
OR severe_red_count >= 2 and score >= 70
```

Where:

- `red_count` is the number of red-level findings in the file;
- `severe_red_count` is the number of red-level findings with high confidence.

This conservative setting avoids marking a file red because of a single statistical anomaly.

### Default finding scores

Approximate default scores:

| Finding type | Default severity | Score |
|---|---:|---:|
| Exact duplicate rows >= 20% | Red, high confidence | 45 |
| Exact duplicate rows >= 5% | Yellow | 18 |
| Normalized duplicate rows >= 10% | Yellow | 20 |
| Duplicate column names | Yellow | 15 |
| Contiguous repeated row run >= 5 | Yellow | 20 |
| Terminal 0/5 share >= 55% | Red | 35 |
| Terminal 0/5 share >= 40% | Yellow | 20 |
| Low terminal-digit entropy | Yellow | 15 |
| Low unique-value ratio with high mode share | Yellow | 18-22 |
| Strong rounding pattern | Yellow | 12 |
| Multiple measurement-like columns tied to row order | Red | 45 |
| Single measurement-like column tied to row order | Yellow | 15 |
| 2D row/column coordinate gradient | Red | 50 |
| Benford first-digit deviation | Yellow | 12 |
| p-values outside `[0, 1]` | Red, high confidence | 50 |
| correlations outside `[-1, 1]` | Red, high confidence | 50 |
| probabilities/AUC outside `[0, 1]` | Red, high confidence | 50 |
| negative count-like values | Red, high confidence | 45 |
| infinite numeric values | Yellow | 20 |
| columns with >=80% missingness | Yellow | 12 |
| all-empty rows >=5% | Yellow | 12 |
| identical byte-level file content | Yellow | 18 |

These scores can be made stricter or more permissive through command-line or GUI threshold settings.

## Common options

```text
--yellow-threshold 25
--red-threshold 90
--min-numeric-n 50
--max-file-mb 250
--max-rows 200000
--max-cols 5000
--sample-rows 100000
```

Example with stricter yellow threshold:

```bash
python raw_data_check_soft/raw_data_audit.py \
  --raw-dir raw_data \
  --yellow-threshold 35 \
  --red-threshold 90
```

## Limitations

- The tool cannot determine scientific validity by itself.
- It does not prove fabrication, fraud, or correctness.
- Some legitimate datasets naturally contain regular structures.
- Very large matrix formats may be audited in metadata-only or sampled form.
- Optional file formats require optional dependencies.
- Report interpretation should involve domain knowledge and inspection of the original experimental design.

## Suggested GitHub contents

Recommended files to upload:

```text
raw_data_check_soft/raw_data_audit.py
raw_data_check_soft/README.md
raw_data_check_soft/requirements.txt
raw_data_check_soft/simulation_data/
```

Recommended files to exclude:

```text
raw_data_check_soft/raw_data_audit_gui.py
raw_data_check_soft/build/
raw_data_check_soft/dist/
raw_data_check_soft/__pycache__/
raw_data_check_soft/*.exe
raw_data_check_soft/*.spec
raw_data_check_soft/build_raw_data_audit_exe.ps1
raw_data_check_soft/raw_data_audit_report.*
raw_data_check_soft/raw_data_audit_simulation_report.*
raw_data_check_soft/~$*.docx
```
