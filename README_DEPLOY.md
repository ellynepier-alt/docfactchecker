# DoC Guideline Fact-Checker Web App

This is a small Flask web app version of the Disorders of Consciousness guideline fact-checker.
It is intended to be hosted on an approved organizational platform, such as an MGB-approved internal hosting environment or an approved cloud tenant.

## What the app does

Users upload a toolkit material and receive a downloadable Word report that compares the material against the AAN-ACRM-NIDILRR Disorders of Consciousness guideline knowledge base.

Supported input types:
- `.txt`
- `.md`
- `.docx`
- `.pdf` with extractable text

The app checks:
- Deprecated or discouraged terminology
- Possible contradictions of the guideline
- Key facts and numeric thresholds
- Evidence-level mismatches
- Coverage map against major recommendations

## Important governance note for MGB

Before putting this on the internet or making it available to participants/users outside your team, route it through the appropriate MGB process. Based on the MGB web hosting guidance you found, a new website requires a Digital Technology Request, and departments should use internal hosting resources where possible. If the website collects, processes, or stores confidential data, additional security/privacy review may be required.

Recommended operating assumptions:
- Do not upload PHI or identifiable patient data.
- Use MGB-approved hosting.
- Use HTTPS only.
- Require access controls if the tool is not meant for public use.
- Add an approved privacy notice if the tool is made public.
- Confirm whether a risk assessment is needed before launch.

## Local testing

1. Install Python 3.10+.
2. From this folder, install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app:

```bash
python app.py
```

4. Open the local URL shown in the terminal, usually:

```text
http://127.0.0.1:5000
```

## Production hosting

This package is intentionally generic. For a real MGB-hosted version, provide this folder to the internal web/application hosting team and ask them to deploy it using MGB-approved hosting and security controls.

Suggested production settings:
- Run behind HTTPS.
- Disable debug mode.
- Set appropriate upload-size limits.
- Store uploads only in temporary storage.
- Configure server logs so they do not retain file contents.
- Add authentication if the tool should be limited to MGB users.

## Files

- `app.py` - Flask web application.
- `factcheck_engine.py` - Guideline checking logic.
- `doc_guidelines_kb.json` - Editable DoC guideline knowledge base.
- `templates/index.html` - Web interface.
- `static/styles.css` - Visual styling.
- `requirements.txt` - Python package dependencies.
