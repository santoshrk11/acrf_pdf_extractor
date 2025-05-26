# aCRF Extractor

Author: Santhosh RK

This project provides a pipeline for extracting metadata and annotations from Annotated Case Report Form (aCRF) PDF documents and producing JSON and Excel reports.

## Project Structure

```
my-acrf-project/
├── data/
│   └── acrf.pdf
├── src/
│   └── extract_acrf.py
├── output/
│   ├── acrf_raw.json
│   ├── acrf_tabular.json
│   └── acrf_report.xlsx
├── .gitignore
├── README.md
└── requirements.txt
```

### Pipeline Flow

```
data/acrf.pdf
    ↓
output/acrf_raw.json
    ↓
output/acrf_tabular.json
    ↓
output/acrf_report.xlsx
```

## Usage

1. Place your non-flattened aCRF PDF in the `data/` directory named `acrf.pdf`.
2. Install the dependencies:

```bash
pip install -r requirements.txt
```

3. Run the extractor script:

```bash
python src/extract_acrf.py
```

The outputs will be generated inside the `output/` directory.

## Requirements

- Python 3.8+
- PyMuPDF 1.22.5
- openpyxl 3.1.2
- python-dateutil 2.8.2

Feel free to extend the parsing logic in `extract_acrf.py` to accommodate different layouts or additional output formats.
