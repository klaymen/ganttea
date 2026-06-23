# Ganttea

Generate a standalone HTML Gantt chart from an Excel spreadsheet.

## Requirements

- Python 3.10+

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python generate_gantt.py test_input.xlsx -o gantt_chart.html
```

Open `gantt_chart.html` in any browser. The file is fully self-contained — no server or internet connection required.

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `input` | `test_input.xlsx` | Path to the input Excel file |
| `-o`, `--output` | `gantt_chart.html` | Output HTML file path |
| `-t`, `--title` | *(derived from input filename)* | Custom chart title |

Example with a custom title:

```bash
python generate_gantt.py test_input.xlsx -o gantt_chart.html -t "Project Timeline Q3"
```

## Input format

The script reads the first sheet of an `.xlsx` file. The header row must include at least these columns (case-insensitive):

| Column | Required | Description |
|--------|----------|-------------|
| Task | Yes | Task or event name |
| Type | Yes | One of `Task`, `Holiday`, or `Milestone` |
| Responsible | No | Person or team name; defaults to `Unassigned` |
| Start Date | Yes | Start date |
| End Date | No | End date; optional for milestones |

Example:

| Task | Type | Responsible | Start Date | End Date |
|------|------|-------------|------------|----------|
| Task-A | Task | John Doe | 2026-06-01 | 2026-06-20 |
| Holiday | Holiday | John Doe | 2026-06-29 | 2026-07-12 |
| Quality Gate | Milestone | | 2026-07-31 | |

Dates can be Excel date cells or strings in `YYYY-MM-DD`, `DD/MM/YYYY`, or `MM/DD/YYYY` format.

## Output

The generated HTML chart includes:

- **Swimlanes** grouped by responsible person
- **Tasks** as blue bars
- **Holidays** as gray striped bars (never shown as milestone markers)
- **Milestones** in a dedicated section at the bottom with diamond markers
- A compact month-based timeline with a legend
- **Dark mode** toggle (remembers your preference in the browser)

## Project files

| File | Description |
|------|-------------|
| `generate_gantt.py` | Main script |
| `test_input.xlsx` | Sample input file |
| `requirements.txt` | Python dependencies |
