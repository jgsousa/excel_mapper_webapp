
# Excel Mapper Web Application (FastAPI + Streamlit)

Guided app to:
1. Upload & analyze **Target.xlsx** (headers row 8, keys from A7 merged range).
2. Upload **sources**, select target **sheets** per source.
3. Run **OpenAI mappings**, **review**, **delete incorrect** ones, and **add missing mappings manually**.
4. **Transfer** (upsert) from sources to target beginning at row 9.

## Manual Mapping (NEW)
In **Step 3**, after you Run/Load the mapping for a `(source → sheet)` pair:
- Review the list (uncheck to remove incorrect entries).
- Use the **Add a missing mapping** controls to pick a **Source header** and a **Target header** from drop-downs.
- Click **Add** to insert it into your working list (we prevent duplicate target columns).
- Click **Save filtered mapping** to persist it server-side.

## Run locally
```
pip install -r requirements.txt
uvicorn backend.main:app --reload --port 8000
streamlit run frontend/streamlit_app.py
```
- UI: http://localhost:8501
- API docs: http://localhost:8000/docs

Set `OPENAI_API_KEY` in your shell or `.env`.

## Docker
Quick start with Docker Compose (two services: API and UI).

### 1) Create `.env`
Copy `.env.example` to `.env` and set your key:
```
OPENAI_API_KEY=sk-...
```

### 2) Build & Run
```
docker compose up --build
```
- Streamlit UI: http://localhost:8501
- FastAPI API: http://localhost:8000/docs

The UI calls the API at `http://api:8000` inside the Compose network. The `./app_data` folder is mounted into both containers for persistence.

### 3) Stop
```
docker compose down
```

### 4) Rebuild after code changes
```
docker compose build --no-cache && docker compose up
```
