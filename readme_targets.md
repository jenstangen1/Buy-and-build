# Methodology: Mapping Potential Targets to B&B Initiatives

This document outlines the methodology implemented in the `map_targets_to_bb.py` script. The primary goal of this script is to identify potential acquisition targets from a master list (`main_target_framework.xlsx`) and map them to relevant Buy & Build (B&B) initiatives based on industry subcategory alignment.

## Input Files

1.  **`main_target_framework.xlsx`**:
    *   **Sheet Name:** `Main`
    *   **Header Row:** 4 (Data starts on row 5)
    *   **Purpose:** Contains a comprehensive list of potential target companies.
    *   **Key Columns Used:**
        *   `Juridisk selskapsnavn`: Target company name.
        *   `NACE-bransjekode`: NACE code for industry classification.
        *   `Sum driftsinnt., 2023`: Revenue figure (assumed to be in NOK thousands).
        *   `Driftsres., 2023`: EBIT figure (assumed to be in NOK thousands).
        *   `Total score`: A score indicating potential exit probability (assumed 0-10).

2.  **`B&B platforms and addons.xlsx`**:
    *   **Sheet Name:** `Data`
    *   **Header Row:** 7 (Data starts on row 8)
    *   **Purpose:** Contains details about existing B&B platforms and add-on acquisitions, including investors.
    *   **Key Columns Used:**
        *   `Companies`: Name of the platform or add-on company.
        *   `All investors` (or similar header containing 'investor'): The private equity firm(s) backing the company.
        *   `Keywords`, `Description`: Text used to classify the company into a segment and subcategory.

## Processing Steps

1.  **Load Target Data:** The script reads the specified columns (`Juridisk selskapsnavn`, `NACE-bransjekode`, `Sum driftsinnt., 2023`, `Driftsres., 2023`, `Total score`) from the `Main` sheet of `main_target_framework.xlsx`. It attempts to map column names robustly, falling back to specific column indices if exact name matches fail for Revenue/EBIT. Revenue and EBIT are loaded as floating-point numbers where possible.

2.  **Load B&B Data:** The script reads the `Data` sheet from `B&B platforms and addons.xlsx`.

3.  **Classify B&B Initiatives:**
    *   It reuses segmentation logic (similar to `segment_classification.py`) based on keywords defined within the script (`segments_data`, `subcategory_keywords`).
    *   Each company from the B&B file is assigned a `Segment` and a `Subcategory` based on matching keywords found in its `Keywords` and `Description` fields.
    *   A lookup dictionary (`subcategory_map`) is created. This map stores, for each `(Segment, Subcategory)` pair, a nested dictionary where keys are investor names. The value for each investor includes a list of their companies (`companies`) in that subcategory and the total count (`count`) of those investments.

4.  **Classify Targets & Filter:**
    *   For each company in the target list, the script extracts the numerical NACE code from the `NACE-bransjekode` column.
    *   It uses a predefined dictionary (`nace_to_subcategory_map`) to map NACE code prefixes (e.g., '41', '43.21', '71.12') to a corresponding `(Segment, Subcategory)` tuple relevant to the Construction & Engineering space.
    *   Targets whose NACE code cannot be mapped to a relevant subcategory or are mapped to the generic "General" subcategory are filtered out.

5.  **Calculate Scores & Format Financials:**
    *   The `Total score` from the target file is converted into an 'Exit Probability (%)' (0-100%), assuming the input score is on a 0-10 scale.
    *   Revenue (`Sum driftsinnt., 2023`) and EBIT (`Driftsres., 2023`) values are formatted into strings like "NOK [number]k", using locale-specific thousands separators. It assumes the input values are already in thousands. Missing or non-numeric financial data is shown as "N/A".

6.  **Match Targets to Potential Acquirers:**
    *   For each *filtered* target, the script uses its assigned `(Segment, Subcategory)` to look up relevant investors and companies in the `subcategory_map` created in step 3.
    *   The result is stored in a 'Potential Acquirers' column, containing the dictionary `{investor: {'companies': [...], 'count': N}}` for that target's subcategory.

7.  **Sort Results:** The final DataFrame containing the filtered and matched targets is sorted by 'Exit Probability (%)' in descending order.

## Output Files

1.  **`potential_targets_overview.html`**: An interactive HTML report displaying the sorted list of potential targets.
    *   Includes columns for Target Name, Segment, Subcategory, Revenue (NOK k), EBIT (NOK k), Exit Probability (%), and Potential Acquirers.
    *   Features filter buttons at the top to view targets by specific subcategories.
    *   The "Potential Acquirers" column lists investors active in the target's subcategory, the number of investments they have in that subcategory, and the names of those companies. For long lists, it shows the first few lines and provides a "Show More..." / "Show Less" toggle.

2.  **`mapped_targets_output.xlsx`**: An Excel file containing the final processed data, including the mapped subcategories, formatted financials, scores, and the raw potential acquirer data. Useful for further analysis or review.

## Key Assumptions & Logic

*   **NACE Mapping:** The core link between targets and B&B initiatives relies on the `nace_to_subcategory_map` dictionary. The accuracy of the mapping depends on this predefined dictionary correctly associating NACE prefixes with the desired subcategories.
*   **B&B Subcategory Classification:** Classification of existing B&B platforms/addons depends on the keyword matching defined in `subcategory_keywords`.
*   **"General" Subcategory:** Targets mapped to a "General" subcategory via NACE codes are excluded from the final report.
*   **Financials:** Revenue and EBIT values in `main_target_framework.xlsx` are assumed to represent **thousands** of NOK.
*   **Exit Score:** The "Total score" column is assumed to be on a scale of 0-10 for calculating the percentage probability. 