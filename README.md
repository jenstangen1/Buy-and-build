# Buy & Build Initiatives Analysis: Methodology Summary

This document outlines the methodologies used in the `segment_classification.py` script to analyze and visualize Buy & Build (B&B) initiatives in the construction sector.

## 1. Segmentation Methodology

Companies from the `B&B platforms and addons.xlsx` file are assigned to broad industry segments based on keyword matching.

-   **Input:** 'Keywords' and 'Description' columns for each company.
-   **Process:**
    1.  A predefined dictionary (`segments_data` in the script) maps segments (e.g., 'Core Construction & Civil Engineering', 'Specialized Trades') to lists of relevant keywords.
    2.  The combined text from a company's 'Keywords' and 'Description' is converted to lowercase.
    3.  This text is searched for matches against the keywords associated with each segment.
    4.  The company is assigned to the *first* segment for which a keyword match is found.
-   **Default:** If no keywords match, the company is assigned to the 'Other' segment.

## 2. Sub-categorization Methodology

For certain segments, companies are further classified into more specific subcategories.

-   **Input:** 'Keywords' and 'Description' columns, and the assigned 'Segment'.
-   **Process:**
    1.  A nested dictionary (`subcategory_keywords` in the script) defines subcategories and associated keywords for specific main segments (e.g., within 'Mechanical, Electrical & HVAC', subcategories like 'HVAC', 'Electrical', 'Plumbing' are defined with their keywords).
    2.  The combined 'Keywords' and 'Description' text is searched for matches against the subcategory keywords *specific to the company's main segment*.
    3.  The company is assigned to the *first* matching subcategory found within its segment.
-   **Default:**
    *   If no specific subcategory keyword matches within the segment's defined subcategories, a default subcategory is assigned (often derived from the main segment name, e.g., 'Core' for 'Core Construction & Civil Engineering').
    *   Segments without predefined subcategories default to a 'General' subcategory.

## 3. Platform / Add-on Classification Methodology

Companies associated with an investor within a specific segment are classified as either 'Platform' or 'Add-on'.

-   **Input:** Company ID, Company Name, Segment, Investor(s), 'Last financing date', 'Revenue'.
-   **Primary Logic (using 'Last financing date'):**
    1.  Companies are sorted by 'Last financing date' (earliest first).
    2.  For each unique combination of `Investor` and `Segment`:
        *   The company with the *earliest* 'Last financing date' is identified.
        *   This earliest company is marked as **'Platform'**.
    3.  All other companies acquired by the *same investor* in the *same segment* are marked as **'Add-on'**.
-   **Revenue Exception:**
    *   After the initial date-based classification, if the 'Revenue' column is available and contains numeric data:
        *   The average revenue for each segment is calculated.
        *   Any company (regardless of acquisition date) with a revenue **more than 5 times** the average revenue of its segment is *also* marked as **'Platform'**. This can potentially reclassify an 'Add-on' to 'Platform' if it's significantly larger than average for its segment.
-   **Fallback Logic (if 'Last financing date' is unavailable):**
    *   If the date column is missing, a simpler classification is used:
        *   For each `Investor` and `Segment` combination, the *first* company encountered in the data processing is marked as **'Platform'**.
        *   All subsequent companies for that investor/segment are marked as **'Add-on'**.
-   **Output:** An `investment_type` ('Platform' or 'Add-on') is assigned to each company in the final `html_data` structure.

## 4. Data Output

The script generates the following output files:

1.  **`bb_initiatives_overview.xlsx`**: An Excel workbook containing:
    *   `Company Classifications`: Raw company data with assigned segments and subcategories.
    *   `B&B Initiatives`: A summary table grouping companies by Investor and Segment.
    *   `Segment Statistics`: Counts of companies per segment.
    *   `Investor Statistics`: Counts of total companies per investor across all segments.
2.  **`construction_segments.csv`**: A CSV file containing the original segment keywords mapping used for classification.
3.  **`bb_initiatives_overview.html`**: An interactive HTML report visualizing the B&B initiatives. It displays companies grouped by segment and investor, includes visual badges for Platform/Add-on status, shows acquisition year, allows filtering by segment/subcategory/investment type, and provides details (Acquired Date, Revenue, EBITDA, Description) on demand. 