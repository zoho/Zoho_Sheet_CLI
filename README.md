# Zoho Sheet CLI — Spreadsheet Intelligence for Agents & Developers

**The fastest way to read, edit, and automate Zoho Sheet spreadsheets — without opening a browser.**

Zoho Sheet CLI is an open-source, cross-platform command-line tool that brings full spreadsheet power to your terminal. Whether you're automating reports, scripting data pipelines, building LLM-powered agents, or just want to edit a cell without loading a GUI, Zoho Sheet CLI gets it done in milliseconds.

---

## Why Zoho Sheet CLI?

- ⚡ **Instant startup** — native Rust binary, no JVM, no interpreter overhead
- 🤖 **Script-friendly** — pipe commands, use in CI/CD, automate anything
- 🖥️ **Works everywhere** — Windows, macOS, Linux (x86 & ARM)
- 📂 **Supports key formats** — `.xlsx`, `.csv`, `.tsv`
- 🔁 **Interactive REPL** — explore and edit spreadsheets live in your terminal
- 🧠 **LLM & agent ready** — clean stdio interface makes it trivial to wire into AI agents, MCP servers, and LLM tool-use workflows
- 🔌 **Powered by Zoho Sheet's native engine** — the same engine behind one of the world's most-used spreadsheet platforms

---

## Quick Start

Install via npm:

```bash
npm install -g @zohocorporation/zs-cli
```

Open a spreadsheet and start editing:

```bash
zs-cli open report.xlsx
```

Or jump into the interactive REPL:

```
$ zs-cli

zs> open mydata.xlsx
✔ Opened: mydata.xlsx (3 sheets)

zs [mydata | Sheet1]> cell get A1
Revenue

zs [mydata | Sheet1]> cell set B2 --formula "=SUM(B3:B100)"
✔ B2 = 42850.00

zs [mydata | Sheet1]> save
✔ Saved: mydata.xlsx
```

## Built for Agentic & LLM Workflows

Zoho Sheet CLI is designed to be a first-class tool in AI-powered pipelines. Its clean, predictable stdio interface means any LLM agent or orchestration framework can drive it without a browser, without a GUI, and without custom integrations.

**Use cases:**
- Give your AI agent the ability to **read and write spreadsheet data** as a tool action
- Wire it into **MCP (Model Context Protocol) servers** to expose spreadsheet operations to Claude, GPT, or any tool-use capable model
- Use it in **LangChain, LlamaIndex, or CrewAI** pipelines to ground agents in real tabular data
- Let LLMs **generate `--script` files** and execute multi-step spreadsheet transformations autonomously

```bash
# An LLM agent can generate and run this in one shot
zs-cli --script agent_generated_pipeline.txt
```

---

## What Can You Do With Zoho Sheet CLI?

| Task | Example |
|------|---------|
| Read a cell | `zs-cli cell get A1` |
| Write a value or formula | `zs-cli cell set B2 --formula "=SUM(B3:B100)"` |
| Export to CSV | `zs-cli save --as output.csv` |
| Run a script of commands | `zs-cli --script pipeline.txt` |
| Manage sheets | `zs-cli worksheet add "Q3 Report"` |
| Find & replace data | `zs-cli replace "Old Co" "New Co"` |
| Sort a data range | `zs-cli sort A1:D100 A` |

---

## Perfect For

- **Data engineers** automating ETL pipelines that touch spreadsheets
- **DevOps teams** generating or validating reports in CI/CD
- **AI/LLM builders** giving agents structured access to spreadsheet data
- **Analysts** who live in the terminal and don't want to open a GUI
- **Developers** building tools that read or write spreadsheet data
- **Anyone** who needs to automate repetitive spreadsheet work

---

## Installation

```bash
npm install -g @zohocorporation/zs-cli
```

> Other package managers (pip, Homebrew, winget) coming soon.

---


