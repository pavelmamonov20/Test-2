# Barcode Excel Generator

This project generates an Excel file with Code128 barcodes (CC001..CC200)
embedded into column B rows.

## Requirements
- Python 3.11+
- Packages listed in `requirements.txt`

## Setup
```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

## Run
```powershell
.venv\Scripts\python generate_barcodes.py
```

## Output
- `barcodes_CC001_to_CC200.xlsx` will be created in the project root.

## Notes
- `.venv/` and generated `.xlsx` files are ignored by Git.
