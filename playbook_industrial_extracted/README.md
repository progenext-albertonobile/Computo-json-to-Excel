# Replit Ready Bundle

This folder is ready to run in Replit as a private/internal app.

## Includes
- `app.py`: Streamlit UI (upload JSON/template, validate, generate XLSX)
- `genera_computo.py`: template-first XLSX engine
- `validate_bundle.py`: CLI validator for capacity/overflow checks
- `requirements.txt`
- `.replit`
- `templates/` directory (put your default template here)

## Setup in Replit
1. Create a new Python Repl.
2. Upload the full `replit_ready` folder content into the project root.
3. Add your default template file at:
   - `templates/Computo preliminare-V3.xlsx`
4. Replit installs dependencies from `requirements.txt`.
5. Click Run.

## Manual run
```bash
pip install -r requirements.txt
streamlit run app.py --server.address 0.0.0.0 --server.port 3000
```

## Validator examples
```bash
python validate_bundle.py test.json
python validate_bundle.py test.json "templates/Computo preliminare-V3.xlsx"
```

## Notes
- Generation is template-driven: styles/formulas/layout stay in the Excel template.
- Dynamic sections/items are supported only within template capacity.
- If overflow is detected, generation is blocked.
