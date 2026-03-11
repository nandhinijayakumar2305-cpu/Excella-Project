# XBP | Dynamic Test Script Generator

**XBP** is a high-performance automation tool designed to bridge the gap between technical field specifications and Quality Assurance workflows.

It intelligently parses **Excel-based data dictionaries** and generates **standardized, audit-ready test scripts**.

---

## 🎯 Key Capabilities

### Heuristic Header Detection
Uses fuzzy matching to identify header rows within complex Excel structures.

### XML-Driven Configuration
Decouples business logic from code, allowing users to map columns via `config.xml` without modifying the Python script.

### Automated Formatting
Applies enterprise-standard styling such as:

- Cell merging
- Color coding
- Font styling
- Structured test case layout

### Validation Logic
Filters and extracts only relevant fields to ensure **clean and accurate test data**.

---

## 🛠 Tech Stack

| Component | Technology |
|----------|-------------|
| Core Logic | Python 3 |
| GUI Framework | Tkinter |
| Data Processing | openpyxl |
| Configuration | XML |

---

## 🚀 Installation

Install dependencies:

```bash
pip install openpyxl
