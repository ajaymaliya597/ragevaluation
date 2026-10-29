# DocRAG Lab — Document Chat + RAG evaluation scaffold

A local FastAPI + HTML/CSS/JS project for uploading documents, asking questions, inspecting retrieved evidence, and exporting an exact audit trail.

## Start (Windows)

1. Install **Python 3.12**.
2. Extract the ZIP.
3. Run `run.bat`. It creates a virtual environment, installs requirements, and starts the app.
4. Open http://127.0.0.1:8000

After the first install, the one-command start can be:
```bat
.venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8000
```
The exact command you requested also works after installing dependencies into the Python 3.12 environment:
```bat
py -3.12 -m uvicorn app:app --host 127.0.0.1 --port 8000
```

## Configure and test Groq (important)

1. Run `run.bat` once, or manually extract the project.
2. In the **same folder as `app.py`**, create a file named exactly `.env` (not `.env.txt`). Copy `.env.example` and fill it like this:

   ```dotenv
   GROQ_API_KEY=gsk_your_real_key_here
   GROQ_MODEL=openai/gpt-oss-20b
   ```

3. Replace the example with your actual key from https://console.groq.com/keys. Do not include `Bearer`, quotes, spaces around the key value, or share/commit the key.
4. Start the app, open http://127.0.0.1:8000, and click **Test Groq connection**. This sends a real tiny request to Groq; `/api/health` only checks whether a key was detected, not whether the key works.
5. If it fails, the UI returns a useful error for missing key, invalid key, unsupported model, quota/rate limit, timeout, or network/firewall failure. See troubleshooting below.

The app reads `.env` on each Groq request, so editing the file normally does not require a restart. Refresh the browser to update the sidebar status. A process-level `GROQ_API_KEY` environment variable takes precedence over `.env`; remove a stale Windows environment variable if it keeps overriding your new key.

### Groq troubleshooting

- **No key detected:** confirm the `.env` file is beside `app.py`; in Windows File Explorer enable **View → Show → File name extensions** and check that it is not `.env.txt`.
- **401/403:** the key is invalid, revoked, or not authorized. Create a fresh key in the Groq Console; do not paste it into chat or screenshots.
- **400/model error:** check that `GROQ_MODEL` names a model available to your Groq account. Keep `openai/gpt-oss-20b` unless Groq reports that it is unavailable.
- **429:** Groq rate limit or quota; wait or check your account limits.
- **Timeout / connection error:** check internet access, VPN/proxy/firewall, and whether `api.groq.com` is reachable from the machine.
- **Port already in use:** close the other server window or use another port and update the browser URL.

To start in the Python 3.12 environment created by `run.bat`:
```bat
.venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8000
```
Your requested command also works when its Python 3.12 environment has all requirements installed:
```bat
py -3.12 -m uvicorn app:app --host 127.0.0.1 --port 8000
```

## Supported formats

PDF, DOCX, XLSX/XLSM, PPTX, CSV, TXT, Markdown, JSON, XML, HTML, YAML, LOG, INI, PY, SQL, and common images. Image OCR requires the Tesseract OCR system executable in addition to Python packages. Scanned PDFs may need OCR preprocessing; normal PDF text extraction is not OCR.

## What gets logged

`data/logs/audit.jsonl` is append-only JSON Lines. Events include timestamps, filenames, SHA-256 hashes, file size, extraction character count, chunk count, request IDs, exact question and answer text, retrieved context, source location, chunk IDs, retrieval scores, model, and latency. Evaluation reports are written to `data/logs/evaluation_<session_id>.json`. Download links are in the UI.

Uploads are stored under `data/uploads/`. Do not upload sensitive documents to an untrusted machine. The app has no authentication and is intended for local development only.

## Live quality evaluation with the actual RAGAS Python library

After each answer, the app runs the official RAGAS Python metric implementations and displays the results below that answer:
- **Faithfulness** — `ragas.metrics.Faithfulness`: whether response claims are supported by retrieved context.
- **Answer relevancy** — `ragas.metrics.AnswerRelevancy` (or the compatible `ResponseRelevancy` name): whether the response addresses the question. This metric uses local Hugging Face embeddings.
- **Context precision** — `ragas.metrics.LLMContextPrecisionWithoutReference`: whether the retrieved chunks are useful for the generated response without requiring a gold reference answer.
- **Evidence coverage** — a separate custom LLM-assisted claim-count audit. It is explicitly *not* an official RAGAS metric; the report labels it separately.

The evaluator LLM uses Groq through LangChain. The embedding model `sentence-transformers/all-MiniLM-L6-v2` runs locally on CPU and is downloaded the first time; first evaluation can therefore take longer and requires internet access to fetch the model. Metric scores are only shown when the actual evaluation returns a score. A failed metric is marked unavailable, never replaced by a guessed percentage. Detailed errors are logged and returned in the evaluation response.

The UI also shows unsupported claims and evaluator notes when available. The “Build evaluation report” button summarizes completed/partial scores across the session.

**Important limitations:** these are model-based estimates of support/relevance, not a guarantee of truth. Faithfulness checks support from retrieved excerpts and cannot prove the original PDF itself is correct. Reference-based context recall, reference-based context precision, and answer correctness/factual correctness require trusted reference answers or labelled ground truth. The app does not fabricate those scores when references are absent.

Some RAGAS metrics need trusted reference answers / labeled examples:
- **Context recall** — needs reference answer/context to determine whether all relevant facts were retrieved.
- **Context precision with reference** — needs a reference answer.
- **Answer correctness / factual correctness** — needs a trusted reference answer or verified labels.
- **Dataset-level RAGAS benchmarking** — needs a curated evaluation dataset; a single live answer cannot replace that benchmark.

For production-grade official RAGAS benchmarking, create a gold test set of questions, reference answers, and (where possible) reference contexts, then run the official RAGAS evaluator with a configured evaluator LLM and compatible embeddings. Do not treat any single aggregate percentage as a guarantee of correctness.

## Retrieval and traceability limitations

The project uses intent-expanded lexical ranking rather than a downloaded semantic embedding model, keeping local setup lightweight. Scores are ranking scores, not calibrated probabilities or semantic similarity. Each chunk is created from a single source evidence block (PDF page, spreadsheet row, slide, paragraph, etc.), preserving source-location fidelity. Scanned PDFs get an OCR fallback when Pillow, pytesseract, and the Tesseract system binary are installed. The retrieved chunk text and IDs are logged exactly so the evidence can be audited. For stronger production traceability, chunk each page/paragraph/table independently and preserve exact page/row ranges.

## API endpoints

- `POST /api/upload` multipart `files`
- `POST /api/chat` JSON `{"session_id":"...","question":"..."}`
- `GET /api/session/{session_id}/logs`
- `POST /api/session/{session_id}/evaluate`
- `GET /api/session/{session_id}/report`
- `GET /api/logs/download`
- `GET /api/health`

Session index is in memory and resets when the server restarts. Uploaded files and audit logs persist.


## Reliability fixes in this build
- Official RAGAS metrics are used for faithfulness, answer relevancy, and context precision; evidence coverage is explicitly marked as a separate custom audit.
- RAGAS is pinned to 0.2.15 for a reproducible metric API, with LangChain Groq and Hugging Face embedding integrations.
- Groq credentials are read from `.env` at request time; the key is never returned by the health endpoint or written to audit logs.
- The UI has a real **Test Groq connection** action that verifies credentials, model access, and network connectivity.
- Groq HTTP errors are translated into actionable messages for invalid credentials, model errors, quota/rate limits, timeouts, and network failures.
- Evidence is chunked by source block and source locations are retained. Scanned PDFs can use an optional OCR fallback.
- Every answer triggers an evidence-based LLM-judge evaluation when Groq is configured; evaluation failures are surfaced rather than replaced with fake scores.
