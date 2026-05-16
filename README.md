# Attendance Report Generator API

FastAPI service for generating professional PDF attendance reports from KIET-format Excel files.

## Endpoints

- `GET /health`
- `POST /reports`
- `/` opens the browser UI

`POST /reports` accepts multipart form data:

- `file`: required `.xlsx` attendance file
- `threshold`: optional attendance threshold, default `75`
- `include_emails`: optional boolean, default `false`

Example:

```bash
curl -X POST "http://localhost:8000/reports" \
  -F "file=@data/1_AI_A_Python.xlsx" \
  -F "threshold=75" \
  -F "include_emails=false" \
  --output attendance-report.pdf
```

## Run locally

```bash
pip install -r requirements.txt
uvicorn api:app --reload
```

Set `OPENAI_API_KEY` before starting the server.

## Deploy on Render

1. Push this repository to GitHub.
2. Create a new Render Web Service from the repository.
3. Add `OPENAI_API_KEY` as an environment variable.
4. Render will use `render.yaml` to install dependencies and start the API.
