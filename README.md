# Excel Case Management Agent

A small Python desktop app that creates and retrieves support cases using an Excel workbook (`cases.xlsx`) as its only data store. You pick a folder, chat with the agent in plain English, and every case is saved to Excel.

By default it uses a built-in **rule-based command parser**, so it needs no API key, network connection, database or paid service. An optional **LLM agent mode** lets a model decide which case action to run by calling tools; it is covered by automated tests using a simulated model. The window header always shows which mode is active.

## Features

- Select any local folder; `cases.xlsx` is created there, or loaded if it already exists (never overwritten).
- Create, list, search (by Case ID) and count cases with natural-language commands.
- Automatic Case IDs (`CASE-0001`, `CASE-0002`, ...) based on the highest existing number, so gaps and deleted rows are handled.
- Excel is the source of truth: records are re-read from the file for every operation.
- Safe saving: write to a temp file, then atomically replace; a `cases.xlsx.bak` backup is made before each change; success is reported only after the record is read back from the saved file.
- Handles missing, corrupt, locked or invalid workbooks, read-only folders and malformed IDs without crashing.
- Formatted workbook: styled bold header, column widths, frozen header row and filters.

## Setup (Windows Command Prompt)

Requires Python 3.11 or 3.12 (the python.org installer includes Tkinter) and Git.

```bat
git clone https://github.com/KeerthanaChandrashekhar/excel-case-agent.git
cd excel-case-agent
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
```

## Run

```bat
python main.py
```

1. Click **Select Folder** and choose any folder.
2. Type a command and press **Enter** (or click **Send**).

## Commands

| Goal | Examples |
|------|----------|
| Create a case | `create case: Payment deducted twice from customer account.` / `Create a case for a customer who cannot access the application.` |
| List all cases | `show all cases`, `list cases`, `display existing cases` |
| Find by ID (case-insensitive) | `find CASE-0003`, `find case-3` |
| Count cases | `how many cases have been created?` |
| Help | `help` |

In rule-based mode these patterns are matched; anything else gets a polite "not recognised" message. In AI mode you can phrase requests freely (for example `my customer can't log in, please log a case`).

### Sample session

```
You:   create case: Payment deducted twice from customer account.
Agent: Case created and saved to Excel.
       Case ID: CASE-0001
       Explanation: Payment deducted twice from customer account.
       Created at: 2026-01-15 09:30:00

You:   how many cases have been created?
Agent: There is 1 case in the workbook.

You:   find case-0001
Agent: Case ID: CASE-0001
       Explanation: Payment deducted twice from customer account.
       Created at: 2026-01-15 09:30:00
```

## Optional: AI agent mode

```bat
set ANTHROPIC_API_KEY=your-key-here
python main.py
```

The header then shows `Mode: AI agent (<model>)`. Set `CASE_AGENT_MODEL` to use a different model. If the key is missing, or the model cannot be reached, the app falls back to the rule-based parser and says so. Never commit your key (`.env` is git-ignored).

## Workbook format

Worksheet `Cases` with columns **Case ID | Explanation | Created At** (`yyyy-mm-dd hh:mm:ss`). You may add or delete rows in Excel while the app is closed; the next ID is always the highest valid suffix + 1.

## Architecture

```
excel-case-agent/
├── main.py            Entry point: logging + launch UI
├── ui.py              Tkinter window (no business logic)
├── agent.py           Routes messages to the AI agent or the rule-based parser
├── llm_agent.py       Optional LLM tool-use loop (create/list/find/count tools)
├── case_manager.py    Validation, ID generation, create/search/list/count
├── excel_handler.py   OpenPyXL I/O, formatting, backup, safe save, errors
├── config.py          Constants (names, headers, patterns, limits)
├── tests/             pytest suite (temporary folders only)
├── requirements.txt
├── pytest.ini
└── .gitignore
```

Flow: `ui` -> `agent.handle(text)` -> (`LLMController` tool calls **or** `parse_command`) -> `CaseManager` -> `ExcelHandler` -> `cases.xlsx`.

**How the AI mode works:** the model is given four tools (create, list, find, count). Each tool runs through `CaseManager`, so Case IDs, timestamps, validation, backups and safe saving are always done by plain Python, never by the model. Tool errors (empty explanation, locked file, failed save) are returned to the model, and the system prompt forbids claiming success unless the tool reported it.

## Tests

```bat
python -m pytest
```

Tests use temporary directories, need no network or API key (the LLM path is tested with a scripted fake model), and never touch your real workbooks.

## Known limitations

- Designed for **one user at a time**; concurrent writes from several people or programs are not supported.
- Close `cases.xlsx` in Excel before creating cases; Windows locks open workbooks and the app will ask you to close it.
- Supported actions are create, list, find and count (no update/close, export/import, categories). The rule-based mode only understands the listed phrasings.
- AI mode needs an internet connection and an Anthropic API key, and sends your messages and the case data returned by tools to the API. It has been tested with a simulated model only, not against the live API.
- Only the most recent backup (`cases.xlsx.bak`) is kept.
- The whole workbook is rewritten on each save: fine for thousands of cases, not for very large datasets.
- Explanations are limited to 32,000 characters; control characters Excel cannot store are removed.

## Possible future work

Update/close cases, export/import, categories and priorities, multi-user support, SAP Joule Studio integration.
