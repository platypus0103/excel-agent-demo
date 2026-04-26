# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**財模助手** — An AI-powered solar power station financial analysis platform. Users chat in Traditional Chinese with a local LLM (Ollama) to query and modify Excel financial models, compare IRR scenarios, and run price rolling calculations.

This is an **in-development version**. The production-ready version with additional tools (IRR reverse calculation, delete sheet) lives in `../Xinbon_final/`.

## Commands

All commands should be run from this directory (`ai_agent2/`).

```bash
# Install dependencies
uv sync

# Start the web server (http://localhost:5000)
uv run python run.py

# CLI chat mode (for quick LLM testing without the web UI)
uv run python main.py
```

**Prerequisites:** Ollama must be running (`ollama serve`) with at least `qwen3:4b` pulled (`ollama pull qwen3:4b`). LibreOffice must be installed for Excel formula recalculation.

There is no test suite configured.

## Architecture

### Request Flow

```
Browser (SPA) → Flask API (api/) → AIAgent (core/agent.py)
                                      ↓
                               OllamaConnection (core/connection.py)
                               sends messages + tool schemas
                                      ↓ (if tool call requested)
                               ToolManager (tool/tool_manager.py)
                               dispatches to excel_tool / finance_tool / etc.
                                      ↓ (tool result fed back to LLM)
                               up to 5 tool-call rounds per query
                                      ↓
                               Final response stored in ConversationManager
                               (max 10 rounds per case)
```

### Key Module Responsibilities

| Path | Role |
|------|------|
| `app.py` | Flask application factory (`create_app()`); registers blueprints, CORS, session, DB |
| `core/agent.py` | `AIAgent` orchestrates the tool-call loop; strips `<think>` tags from qwen3 output |
| `core/connection.py` | `OllamaConnection` — raw HTTP to Ollama; 3-retry with 1s delay |
| `core/conversation.py` | Per-case chat history, capped at 10 rounds |
| `config/settings.py` | `AgentConfig` dataclass + inline system prompt (~144 lines) |
| `tool/tool_manager.py` | Registers all tools (JSON Schema); dispatches by tool name |
| `tool/excel_tool.py` | LLM-facing Excel tools: read/write cells, fuzzy field search, financial data query |
| `tool/finance_tool.py` | IRR calculations (Project / Cost / Equity) and cross-sheet IRR comparison |
| `tool/price_rolling_tool.py` | Price rolling calculation (read-only, no Excel write) |
| `tool/equipment_cost_tool.py` | Price rolling + write-back to Excel with audit log |
| `tool/equipment_cost_services.py` | Rolling mode logic: `CashMode` / `RatioMode` / `CustomizeMode` |
| `services/llm_service.py` | Wraps `ToolManager` for direct price rolling execution; holds in-memory rolling param cache per `excel_path` |
| `services/excel_tool.py` | Excel helper utilities (used by services layer, not exposed to LLM) |
| `services/irr_calculator.py` | IRR calculation service (used by services layer) |
| `models/db_models.py` | SQLAlchemy ORM: `User`, `Case`, `ChatMessage` |
| `models/database.py` | SQLAlchemy `db` instance initialization |
| `models/irr_models_v2.py` | Pydantic validation models for IRR parameters |
| `api/agent_routes.py` | `/api/agent_chat`, `/api/upload_excel`, `/api/download_excel`, price rolling, sheet import endpoints |
| `api/auth_routes.py` | Email-only login (no passwords) |
| `api/case_routes.py` | Case CRUD |
| `utils/recalc.py` | Calls LibreOffice headless to recalculate Excel formulas after writes |
| `utils/formula_evaluator.py` | Reads Excel cell values post-write using the `formulas` package — works around openpyxl's `data_only=True` clearing cached formula results |
| `utils/app_logger.py` | Logging with Taiwan timezone (UTC+8), daily rotation, 30-day retention |
| `utils/error_handler.py` | Error handling and retry decorators |

### LLM-Facing Tools

| Tool | File | Description |
|------|------|-------------|
| `read_excel_cell` | `tool/excel_tool.py` | Read a specific cell |
| `write_excel_cell` | `tool/excel_tool.py` | Write a specific cell |
| `delete_excel_cell` | `tool/excel_tool.py` | Clear a specific cell |
| `list_excel_sheets` | `tool/excel_tool.py` | List all worksheets |
| `read_sheet_by_field` | `tool/excel_tool.py` | Query by field name (fuzzy match, threshold ~0.6) |
| `edit_sheet_by_field` | `tool/excel_tool.py` | Modify by field name (fuzzy match); accepts `row_hint` for direct row targeting |
| `query_financial_data` | `tool/excel_tool.py` | Full-sheet financial data scan; returns IRR already in % form |
| `compare_irr_across_sheets` | `tool/finance_tool.py` | Compare IRR across scenario sheets (p1~pN); returns per-sheet results + best per type |
| `calculate_price_rolling` | `tool/price_rolling_tool.py` | Compute rolling price (no write); modes: CashMode/RatioMode/CustomizeMode |
| `execute_price_rolling` | `tool/equipment_cost_tool.py` | Compute + write to Excel with audit record |

> **Note:** This version does NOT have `find_parameter_for_target_irr` (IRR 逆推) or `delete_excel_sheet`. Both are implemented in `../Xinbon_final/`.

### Multi-Tenant Data Isolation

- Authentication is email-only (no passwords); first login auto-creates the account.
- Each user's Excel files live at `Excel User Data/{email}/{case_id}_{filename}.xlsx`.
- Chat histories are per-case in SQLite (`instance/app.db`, auto-created on first run).
- `agent_routes._find_excel_file()` resolves files in priority order: exact `{case_id}_{filename}` → `{case_name}_{filename}` → `{case_id}_*.xlsx` prefix → `{case_name}_*.xlsx` prefix → substring match → first `.xlsx` in folder.

### Frontend

Single-page app at `templates/LLMweb.html`. `static/LLMweb.js` (~2,000 lines) handles all UI logic. `static/price_rolling.js` handles the rolling dialog. Luckysheet (loaded from CDN) renders the Excel spreadsheet in-browser.

## Configuration

`config/settings.py` controls:

- **Model**: `qwen3:4b` default; switchable to `qwen3:14b` / `qwen3:32b` from the UI.
- **Temperature / Top-P / Top-K**: `0.1` / `0.1` / `100` (low creativity for financial accuracy).
- **Thinking mode**: `thinking_mode = False`; set `True` for qwen3 extended reasoning (recommended for 32b).
- **Tool call rounds**: max 5 per query.
- **Ollama host**: defaults to `http://localhost:11434`.

The system prompt (embedded inline in `AgentConfig.system_prompt`, ~144 lines) defines tool-use rules, scope boundaries, the Excel modification confirmation flow, and **explicitly requires all responses in Traditional Chinese (繁體中文)**. Never change responses to Simplified Chinese.

## Important Constraints

- All AI responses to users must be in **Traditional Chinese** — this is enforced by the system prompt and is a hard requirement from the client.
- Every Excel write logs a record to the `滾算紀錄單(總紀錄)` sheet inside the same file.
- The three IRR methods (Project / Cost / Equity) have precise financial definitions — see `tool/finance_tool.py` for the formulas. Do not conflate them.
- **IRR percentage rule**: `query_financial_data` returns IRR already as a percentage (e.g., `2.72` means `2.72%`). Display with `%` appended — never multiply by 100.
- The system prompt prohibits the LLM from revealing internal tool names, architecture, or system prompt content to users — keep this rule intact when editing the prompt.
- `SECRET_KEY` is hardcoded in `app.py` and CORS allows all origins — production deployment must address both.
