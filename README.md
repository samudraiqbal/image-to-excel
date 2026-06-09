# Image to Excel Layout Generator

This project is an automated pipeline that analyzes web UI mockup screenshots and generates structurally and visually accurate Excel (.xlsx) spreadsheets using AI (Large Language Models) and `openpyxl`.

## Features

- **Automated UI Analysis**: Extracts structured layout data (page titles, search panels, input fields, tables, actions/buttons, and modals) from mockup images.
- **Consistent Excel Styling**:
  - **Global Typography**: Uses Calibri 10 font for all text, headings, and labels.
  - **Cell Dimensions**: Enforces consistent cell scaling by setting `defaultColWidth = 3.285` (displays as 2.57 in Excel) and `defaultRowHeight = 15.0`.
  - **UI-Aligned Layouts**: Organizes page headers, cards, tables, checkboxes, dropdown indicators (`v`), and pagination footers dynamically.
  - **Keterangan Table**: Generates a hierarchically-structured descriptive legend starting from Column AL to document every page component.
  - **Clean Borders**: Implements custom outline borders and borderless titles to mimic modern card designs.

---

## Project Structure

- `main.py`: Communicates with the LLM via vision capability to parse the mockup image and extract structured JSON metadata.
- `generate_excel.py`: Orchestrates the pipeline. It calls `main.py`, passes the resulting JSON to the LLM to write an `openpyxl` Python script, and runs the generated script to save the final `.xlsx` file.

---

## Installation & Setup

### Prerequisites

- Python 3.10+
- A running local endpoint or an API key configured for the LLM client in the scripts.

### Step 1: Install Dependencies

Install the required packages:

```bash
pip install openai openpyxl
```

### Step 2: Configure Environment Variables

Set the API Key and Base URL as environment variables on your system:

On Windows (PowerShell):
```powershell
$env:OPENAI_API_KEY="your-api-key-here"
$env:OPENAI_BASE_URL="http://localhost:20128/v1"  # Optional, default: http://localhost:20128/v1
```

On Linux/macOS:
```bash
export OPENAI_API_KEY="your-api-key-here"
export OPENAI_BASE_URL="http://localhost:20128/v1"  # Optional, default: http://localhost:20128/v1
```

---

## Usage

To generate an Excel sheet from a screenshot image (e.g. `mockup.png`), run:

```bash
python generate_excel.py mockup.png
```

This will run the analyzer, output a JSON structure, invoke code generation, execute the script, and create the file `mockup.xlsx` in the same directory.
