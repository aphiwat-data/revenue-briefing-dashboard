# G5 D4cast Revenue Dashboard

## Deploy target
This package is ready for Streamlit Community Cloud.

## Files
- `app.py` — Streamlit dashboard entrypoint
- `requirements.txt` — Python dependencies

## Important
For Streamlit Cloud deployment, local Windows paths like `G:\My Drive\...` will not exist on the cloud server. Use **Manual upload** for the presentation/demo version.

If you want true automatic Google Drive loading on cloud, the next step is Google Drive API + service account / OAuth integration.

## Local run
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Streamlit Cloud setup
Entrypoint file:
```text
app.py
```
