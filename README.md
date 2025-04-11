# PubMed to XLSX Training Data Extractor

This Python script automates the creation of training data for a machine learning model developed to assist the **Online Mendelian Inheritance in Animals (OMIA)** project.

## 🔍 Project Purpose

OMIA ([omia.org](https://omia.org)) is an online catalogue of inherited disorders in animals. To maintain and grow this catalogue, OMIA regularly adds references to new, relevant scientific papers from PubMed. A machine learning model is being developed to **shortlist likely-relevant papers** from the daily influx of biomedical literature. This script helps prepare **training data** for that model.

## 🧠 What This Script Does

- Takes a plain text `.txt` file of new PubMed citations for one or more days in **MEDLINE format**
- Parses each paper entry, separated by empty lines
- Extracts the following fields:
  - `PMID` — PubMed ID
  - `TI` — Title
  - `AB` — Abstract
- Outputs the data to an `.xlsx` spreadsheet with columns:
  - **PMID**, **TI**, **AB**

### 📄 Example Input (MEDLINE format)

```text
PMID- 12345678
TI  - Title of the paper.
AB  - Abstract text goes here.
AU  - Smith J
DP  - 2025 Apr 10

PMID- 12345679
TI  - Another paper title.
AB  - Another abstract.
AU  - Doe A
DP  - 2025 Apr 10
```

### 📦 Output Format

The script produces an `.xlsx` file like this:

| PMID      | TI                     | AB                    |
|-----------|------------------------|------------------------|
| 12345678  | Title of the paper.    | Abstract text goes here. |
| 12345679  | Another paper title.   | Another abstract.         |

## 🚀 Usage

1. Place your PubMed `.txt` file (in MEDLINE format) in the '/original_pubmed_text' directory.
2. Run the script:
   ```bash
   python pmparser.py
   ```
3. The script will generate an Excel file in ./processed_xlsx/ with the extracted training data.

## 📦 Requirements

- Python 3.7+
- `xlsxwriter` (install via `conda install xlsxwriter`)

## 🧪 Notes

- The script handles missing abstracts or titles by filling in blanks.
- Input files must be in MEDLINE text format.

