# DoC Guideline Fact-Checker, Streamlit Version

This is a web-based Streamlit version of the Disorders of Consciousness guideline fact-checker.

It lets someone:
1. Open a browser-based interface.
2. Upload a toolkit material.
3. Run a guideline alignment check.
4. Download a Word report.

## Files included

- `app.py` - Streamlit web app.
- `factcheck_engine.py` - Fact-checking logic.
- `doc_guidelines_kb.json` - Editable guideline knowledge base.
- `requirements.txt` - Python dependencies.
- `.streamlit/config.toml` - Visual theme.

## Run locally on a Mac

Open Terminal and go to the folder that contains these files.

Install the required packages:

```bash
python3 -m pip install -r requirements.txt
```

Run the app:

```bash
python3 -m streamlit run app.py
```

A browser window should open automatically. If not, Terminal will show a local URL that starts with `http://localhost`.

## Run locally on Windows

Open Command Prompt and go to the folder that contains these files.

Install the required packages:

```bash
py -m pip install -r requirements.txt
```

Run the app:

```bash
py -m streamlit run app.py
```

## Deploy to a web-based platform

This package can be deployed to a Streamlit-compatible hosting environment. For MGB use, use an MGB-approved environment and route through the appropriate Digital Technology / security / privacy process before making it available broadly.

Do not host this publicly with sensitive or identifiable data. The intended use is educational toolkit review, not patient care or clinical decision support.

## Customize the guideline rules

Open `doc_guidelines_kb.json` and edit:

- `terminology_flags`
- `contradiction_flags`
- `recommendations`

Save the file and rerun the app.
