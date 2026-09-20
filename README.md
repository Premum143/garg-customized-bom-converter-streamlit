# GARG Customized BOM Converter

Streamlit app for converting GARG customized BOM workbooks into TranZact assembly BOM format.

## Scope

- Assembly BOM only
- Cut parts are assumed to already exist in inventory
- Cut-part BOM and nesting are handled outside this tool
- Supports:
  - flat BOMs where one FG directly consumes cut parts
  - section-based multi-level BOMs where `GEL_..._A01/A02/...` rows define SFG sections

## Output Sheets

- `FG`
- `RM`
- `Routing`
- `Master File`

## Item Logic

- FG Item ID = assembly drawing number
- FG Item Name = BOM description
- Cut-part Item ID = `DRAWING NO`
- Cut-part Item Name = `DESCRIPTION + MATERIAL CODE + size`

## Default Stores

- FG Store: `Garg Unit - 1`
- RM Store: `Garg Unit - 2`
- Reject Store: `Default Reject Store`
- WIP Store: `Default WIP Store`

## Local Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

Use:

- Repository: this repo
- Branch: `main`
- Main file path: `app.py`

The app is intentionally open and does not include password protection.
