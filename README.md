# Buy & Build Analysis Tool

A Python-based tool for analyzing buy-and-build initiatives in the construction and engineering sector, with a focus on Norway and Sweden.

## Features

- Processes Excel data of companies and investors
- Classifies companies into industry segments and subcategories
- Identifies platform investments and add-ons
- Generates an interactive HTML report with:
  - Company details and financial metrics
  - Segment-based filtering
  - Platform/Add-on classification
  - Subcategory filtering
  - Investor grouping

## Setup

1. Ensure you have Python 3.x installed
2. Install required packages:
   ```bash
   pip install pandas numpy openpyxl
   ```

## Usage

1. Place your Excel files in the project directory:
   - `B&B platforms and addons.xlsx`
   - `B&B investors.xlsx`

2. Run the analysis:
   ```bash
   python segment_classification.py
   ```

3. View the results:
   - Open `bb_initiatives_overview.html` in a web browser
   - Check `bb_initiatives_overview.xlsx` for detailed data

## Output Files

- `bb_initiatives_overview.html`: Interactive visualization of all data
- `bb_initiatives_overview.xlsx`: Detailed Excel report with multiple sheets
  - Company Classifications
  - B&B Initiatives
  - Segment Statistics
  - Investor Statistics 