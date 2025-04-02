# map_targets_to_bb.py

import pandas as pd
import numpy as np
import re
import datetime
import math
import locale # For number formatting

# Set locale for number formatting (e.g., Norwegian thousands separator)
try:
    # Try setting to Norwegian Bokmål for typical separators
    locale.setlocale(locale.LC_ALL, 'nb_NO.UTF-8')
except locale.Error:
    try:
        # Fallback to a generic English locale if Norwegian isn't available
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
    except locale.Error:
        # Fallback if no locale setting works
        print("Warning: Could not set locale for number formatting.")

# --- Reuse Subcategory Logic from segment_classification.py ---

# Define the segments and their associated keywords (copied for context)
segments_data = {
    'Core Construction & Civil Engineering': [
        'construction company', 'construction work', 'construction services',
        'construction contractor', 'construction operation', 'construction project',
        'contracting services', 'excavation work', 'groundwork service',
        'demolition services', 'road construction', 'tunneling', 'rock blasting',
        'builder services', 'building developer', 'house builder',
        'house construction', 'residential construction', 'modular buildings',
        'real estate development', 'property development', 'property maintenance',
        'property management', 'property renovation'
    ],
    'Specialized Trades': [
        'flooring services', 'floor leveling', 'tiling', 'tiles work',
        'tiles laying', 'flooring company', 'floor treatment',
        'carpentry work', 'carpentry products', 'painting services',
        'painting provider', 'wall tiles', 'surface layering',
        'roofing services', 'roofing system', 'roofing maintenance',
        'pitched roof', 'flat roof'
    ],
    'Mechanical, Electrical & HVAC': [
        'electrical installation', 'electrical engineering', 'electrical contractor',
        'power installation', 'lighting systems', 'hvac services', 'heating system',
        'ventilation system', 'air-treatment installation', 'indoor climate',
        'plumbing services', 'pipe installation', 'drainage systems',
        'geothermal heating'
    ],
    'Marine, Offshore & Energy': [
        'diving & salvage', 'marine construction', 'marine survey',
        'hydrographic survey', 'offshore unit', 'oil and gas support',
        'oil and gas pipe', 'oil and gas investment', 'solar energy',
        'energy efficiency', 'renewable energy', 'energy consulting'
    ],
    'Industrial Services & Manufacturing Support': [
        'welding services', 'steel cutting', 'forging', 'machining',
        'rotating equipment maintenance', 'concrete pumping', 'precast concrete',
        'foam concrete', 'concrete technology', 'concrete renovation',
        'machine control', 'process plant maintenance', 'repair and maintenance'
    ],
    'Building Products & Materials': [
        'building materials', 'modular walls', 'glass walls',
        'fire retardant wood', 'surface materials', 'insulation',
        'prefabricated housing', 'green construction materials',
        'constructional material'
    ],
    'Tech & Software for Construction': [
        'project management system', 'construction management software',
        'online collaboration', 'bim', '3d modeling', 'reverse engineering',
        'smart building systems'
    ],
    'Consulting, Advisory & Project Management': [
        'construction consulting', 'engineering consulting', 'financial advisory',
        'spatial planning', 'architectural consultancy', 'design management',
        'project planning', 'cost estimation', 'technical consulting',
        'environmental consulting', 'geotechnical consulting', 'remediation services'
    ],
    'Equipment Rental & Heavy Machinery': [
        'machinery rental', 'construction equipment', 'crane trucks',
        'scaffolding', 'automation machinery', 'heavy equipment services',
        'pipeline services', 'construction machines repair'
    ],
    'Facility Services & Real Estate Ops': [
        'building automation', 'smart buildings', 'energy management systems',
        'property tech', 'maintenance services', 'damage restoration',
        'dehumidification', 'climate control', 'insurance claims',
        'remediation', 'fire & water damage control'
    ],
    'Safety & Monitoring Systems': [
        'alarm systems', 'surveillance', 'access control', 'fire safety',
        'radon mitigation', 'system installations'
    ],
    'Environmental & Waste Management': [
        'waste management', 'land remediation', 'environmental services',
        'asbestos removal', 'soil mixing', 'decontamination'
    ],
    'Infrastructure & Public Works': [
        'infrastructure construction', 'road development', 'railway development',
        'traffic systems', 'public facility works'
    ],
    'Interior Design & Furnishing': [
        'interior services', 'office decor', 'store design',
        'modular interiors', 'window and furnishing installation'
    ]
}

subcategory_keywords = {
    'Mechanical, Electrical & HVAC': {
        'HVAC': ['hvac', 'ventilation', 'air', 'climate', 'cooling', 'varme', 'kjøling'],
        'Electrical': ['electrical', 'power', 'lighting', 'elektro'],
        'Plumbing': ['plumbing', 'pipe', 'drainage', 'sanitary', 'rørlegger'],
        'Heating': ['heating', 'geothermal', 'heat pump', 'boiler', 'oppvarming']
    },
    'Core Construction & Civil Engineering': {
        'Civil Engineering': ['civil engineering', 'infrastructure', 'road', 'tunnel', 'excavation', 'groundwork', 'anlegg', 'vei', 'grunnarbeid'],
        'Building Construction': ['building', 'construction company', 'house builder', 'residential', 'bygge', 'entreprenør'],
        'Property Services': ['property', 'real estate', 'renovation', 'maintenance', 'eiendom', 'rehabilitering', 'vedlikehold']
    },
    'Specialized Trades': {
        'Flooring': ['floor', 'tiling', 'tiles', 'gulv', 'flis'],
        'Carpentry': ['carpentry', 'wood', 'timber', 'tømrer', 'snekker'],
        'Roofing': ['roof', 'roofing', 'tak'],
        'Painting': ['paint', 'coating', 'surface', 'male', 'overflate']
    },
    'Industrial Services & Manufacturing Support': {
        'Welding & Metalwork': ['welding', 'steel', 'metal', 'forging', 'sveising', 'stål', 'metall'],
        'Concrete': ['concrete', 'cement', 'betong'],
        'Maintenance': ['maintenance', 'repair', 'service', 'vedlikehold', 'reparasjon']
    }
    # NOTE: Other segments need a 'General' subcategory
}

# Add a default subcategory for all other segments
for segment in segments_data:
    if segment not in subcategory_keywords:
        # Ensure the main segment key exists
        subcategory_keywords[segment] = {}
    if 'General' not in subcategory_keywords[segment]:
         # Add 'General' if not already defined (prevents overwriting existing subcats)
         # Use the first word of the segment name as a keyword for 'General'
        general_keyword = segment.split(' ')[0].lower()
        subcategory_keywords[segment]['General'] = [general_keyword]


# Function to find matching segment based on keywords and description
def find_matching_segment(row, segments_data):
    text_to_search = f"{str(row.get('Keywords', '')).lower()} {str(row.get('Description', '')).lower()}"
    matching_segments = []

    for segment, keywords in segments_data.items():
        for keyword in keywords:
            # Basic check if keyword exists as a whole word or part of the text
            if re.search(r'\b' + re.escape(keyword.lower()) + r'\b', text_to_search):
                 matching_segments.append(segment)
                 break # Match found in this segment, move to next segment

    if matching_segments:
        # Prioritize more specific segments if multiple matches? For now, take first.
        return matching_segments[0]
    return "Other" # Default if no keywords match


# Function to identify subcategory based on text
def identify_subcategory(text, segment):
    # Default subcategory is 'General' within the segment
    default_subcategory = 'General'

    if not text or not isinstance(text, str):
        return default_subcategory

    text_lower = text.lower()

    # Get subcategory definitions for this segment
    if segment in subcategory_keywords:
        subcategories = subcategory_keywords[segment]
        # Check specific subcategories first (before 'General')
        specific_subcats = {k: v for k, v in subcategories.items() if k != 'General'}

        for subcategory, keywords in specific_subcats.items():
            for keyword in keywords:
                 # Use regex for potentially more robust matching (whole words)
                if re.search(r'\b' + re.escape(keyword.lower()) + r'\b', text_lower):
                    return subcategory # Return specific subcategory if keyword matches

    # If no specific subcategory found, return 'General'
    return default_subcategory

# --- NACE Code to Subcategory Mapping ---
# Mapping NACE prefixes to (Segment, Subcategory) tuples
# Based on NACE Rev. 2 and subcategories defined above
nace_to_subcategory_map = {
    # F - CONSTRUCTION
    '41': ('Core Construction & Civil Engineering', 'Building Construction'),
    '42': ('Core Construction & Civil Engineering', 'Civil Engineering'),
    '43.1': ('Core Construction & Civil Engineering', 'Civil Engineering'), # Demolition, site prep
    '43.21': ('Mechanical, Electrical & HVAC', 'Electrical'),
    '43.22': ('Mechanical, Electrical & HVAC', 'HVAC'), # Plumbing, heat, air-con
    '43.29': ('Specialized Trades', 'General'), # Other installation
    '43.31': ('Specialized Trades', 'Painting'), # Plastering often related
    '43.32': ('Specialized Trades', 'Carpentry'), # Joinery
    '43.33': ('Specialized Trades', 'Flooring'), # Floor and wall covering
    '43.34': ('Specialized Trades', 'Painting'), # Painting and glazing
    '43.39': ('Specialized Trades', 'General'), # Other completion
    '43.91': ('Specialized Trades', 'Roofing'),
    '43.99': ('Specialized Trades', 'General'), # Other specialized activities (scaffolding etc.)

    # M - PROFESSIONAL, SCIENTIFIC AND TECHNICAL ACTIVITIES
    '71.11': ('Consulting, Advisory & Project Management', 'General'), # Architectural activities
    '71.12': ('Consulting, Advisory & Project Management', 'General'), # Engineering activities & consultancy
    '71.2': ('Consulting, Advisory & Project Management', 'General'), # Technical testing and analysis

    # N - ADMINISTRATIVE AND SUPPORT SERVICE ACTIVITIES
    '77.32': ('Equipment Rental & Heavy Machinery', 'General'), # Renting construction machinery
    '81.1': ('Facility Services & Real Estate Ops', 'General'), # Combined facilities support
    '81.29': ('Facility Services & Real Estate Ops', 'General'), # Other building and industrial cleaning (can include exterior etc)
    '81.3': ('Facility Services & Real Estate Ops', 'General'), # Landscape service activities

    # C - MANUFACTURING
    '23.5': ('Building Products & Materials', 'General'), # Cement, lime, plaster
    '23.6': ('Industrial Services & Manufacturing Support', 'Concrete'), # Concrete, plaster, cement products
    '23.7': ('Building Products & Materials', 'General'), # Cutting, shaping stone
    '25.1': ('Industrial Services & Manufacturing Support', 'Welding & Metalwork'), # Structural metal products
    '32.99': ('Building Products & Materials', 'General'), # Other manufacturing n.e.c. (can include signs etc)

    # E - WATER SUPPLY; SEWERAGE, WASTE MANAGEMENT AND REMEDIATION ACTIVITIES
    '37': ('Core Construction & Civil Engineering', 'Civil Engineering'), # Sewerage (often infrastructure related)
    '38': ('Environmental & Waste Management', 'General'), # Waste collection, treatment
    '39': ('Environmental & Waste Management', 'General'), # Remediation activities

    # G - WHOLESALE TRADE
    '46.73': ('Building Products & Materials', 'General'), # Wholesale of wood, construction materials
    '46.74': ('Building Products & Materials', 'General'), # Wholesale of hardware, plumbing and heating equipment
}

def get_subcategory_from_nace(nace_code_text):
    """Extracts NACE code and maps it to a (Segment, Subcategory) tuple."""
    if pd.isna(nace_code_text):
        return None, None

    # Extract numeric code (e.g., '41.201' or '41201' from '41.201 - Bygging av bygninger')
    match = re.match(r'^\s*(\d{2}(\.\d{1,3})?|\d{4,5})', str(nace_code_text))
    if not match:
        return None, None # No valid code found at the start

    nace_code = match.group(1).replace('.', '') # Standardize to remove dots initially

    # Try matching prefixes in order of decreasing specificity
    # Check for 5-digit match (e.g., 43.21x) -> 4-digit (e.g., 43.2x) -> 3-digit (e.g., 43.x) -> 2-digit (e.g., 4x)
    # Re-insert dots for matching keys in nace_to_subcategory_map
    prefixes_to_try = []
    if len(nace_code) >= 4:
        prefixes_to_try.append(f"{nace_code[:2]}.{nace_code[2:4]}") # e.g. 43.21
    if len(nace_code) >= 3:
         prefixes_to_try.append(f"{nace_code[:2]}.{nace_code[2]}")  # e.g. 43.1
    if len(nace_code) >= 2:
        prefixes_to_try.append(nace_code[:2]) # e.g. 41, 42, 43

    for prefix in prefixes_to_try:
         if prefix in nace_to_subcategory_map:
              return nace_to_subcategory_map[prefix] # Return (Segment, Subcategory)

    return None, None # No relevant mapping found

# --- Formatting Function --- (Revised)
def format_nok_thousands(value):
    """Formats a number (assumed to be in thousands) as NOK thousands string."""
    if pd.isna(value):
        return 'N/A'
    try:
        # Value is already in thousands, just format it
        num = float(value)
        formatted_num = locale.format_string("%.0f", num, grouping=True)
        return f"NOK {formatted_num}k"
    except (ValueError, TypeError):
        # Handle cases where conversion to float fails
        return 'N/A' # Changed from 'Invalid' to 'N/A'

# --- Main Script Logic ---
if __name__ == "__main__":
    print("Starting target mapping process...")

    # 1. Load Target Data (including Revenue and EBIT)
    target_file = 'main_target_framework.xlsx'
    target_sheet = 'Main'
    print(f"\nLoading target data from '{target_file}' sheet '{target_sheet}'...")
    try:
        df_target_headers = pd.read_excel(target_file, header=3, sheet_name=target_sheet, nrows=0)
        target_cols_map = {}
        rev_col_found = False
        ebit_col_found = False

        # Find actual column names robustly using the provided exact names
        for col in df_target_headers.columns:
            col_str = str(col) # Ensure it's a string
            col_lower = col_str.lower()

            if 'juridisk selskapsnavn' in col_lower: target_cols_map['TargetName'] = col_str
            elif 'nace-bransjekode' in col_lower: target_cols_map['NACE'] = col_str
            elif 'total score' in col_lower: target_cols_map['Score'] = col_str
            # Use exact names provided by user (case-insensitive check for safety)
            elif col_lower == 'sum driftsinnt., 2023'.lower():
                 target_cols_map['Revenue'] = col_str
                 rev_col_found = True
            elif col_lower == 'driftsres., 2023'.lower():
                 target_cols_map['EBIT'] = col_str
                 ebit_col_found = True

        required_cols = ['TargetName', 'NACE', 'Score'] # Base requirements
        # Add Rev/EBIT to required if found
        if rev_col_found: required_cols.append('Revenue')
        else: print("Warning: Revenue column 'Sum driftsinnt., 2023' not found.")
        if ebit_col_found: required_cols.append('EBIT')
        else: print("Warning: EBIT column 'Driftsres., 2023' not found.")


        # Check if all *found/expected* required columns are in the map
        # We proceed even if Rev/EBIT not found, but base columns must exist
        base_required = ['TargetName', 'NACE', 'Score']
        if not all(col in target_cols_map for col in base_required):
             missing = [col for col in base_required if col not in target_cols_map]
             raise ValueError(f"Missing essential columns in target file: {missing}. Found headers: {list(df_target_headers.columns)}")

        # Specify data types for Rev/EBIT during load if possible
        dtype_spec = {}
        if 'Revenue' in target_cols_map: dtype_spec[target_cols_map['Revenue']] = float
        if 'EBIT' in target_cols_map: dtype_spec[target_cols_map['EBIT']] = float

        # Read only the columns we actually found
        cols_to_read = [target_cols_map[req] for req in required_cols if req in target_cols_map]

        df_target = pd.read_excel(
            target_file,
            header=3,
            sheet_name=target_sheet,
            usecols=cols_to_read, # Use the found column names
            dtype=dtype_spec if dtype_spec else None # Apply dtypes
        )
        # Rename columns to standardized names based on what was found
        rename_map = {v: k for k, v in target_cols_map.items() if v in cols_to_read}
        df_target = df_target.rename(columns=rename_map)

        # Ensure Revenue and EBIT columns exist even if mapping failed (fill with NaN)
        if 'Revenue' not in df_target.columns: df_target['Revenue'] = np.nan
        if 'EBIT' not in df_target.columns: df_target['EBIT'] = np.nan

        print(f"Loaded {len(df_target)} target companies.")
        print("Target columns found and mapped:", target_cols_map)

    except FileNotFoundError: print(f"Error: Target file '{target_file}' not found."); exit()
    except ValueError as e: print(f"Error loading target data: {e}"); exit()
    except Exception as e: print(f"An unexpected error occurred loading target data: {e}"); exit()

    # 2. Load and Process B&B Platform/Addon Data
    bb_file = 'B&B platforms and addons.xlsx'
    print(f"\nLoading B&B data from '{bb_file}'...")
    try:
        df_platforms = pd.read_excel(bb_file, header=6, engine='openpyxl', sheet_name='Data')
        df_platforms = df_platforms.dropna(how='all').dropna(axis=1, how='all')

        # Find 'All investors' column more robustly
        investor_col_name = None
        for col in df_platforms.columns:
             col_str = str(col)
             if 'investor' in col_str.lower():
                  investor_col_name = col_str
                  break
        if not investor_col_name:
             # Default if not found, but might cause issues later if column truly missing
             investor_col_name = 'All investors'
             if investor_col_name not in df_platforms.columns:
                  print(f"Warning: Could not find 'All investors' column in {bb_file}.")
                  df_platforms[investor_col_name] = '' # Add empty column to prevent key errors

        # Ensure base required columns exist
        base_req = ['Companies', 'Keywords', 'Description']
        if not all(c in df_platforms.columns for c in base_req):
             raise ValueError(f"Missing base columns in {bb_file}: {[c for c in base_req if c not in df_platforms.columns]}")


        print(f"Loaded {len(df_platforms)} platform/addon records.")
        # Classify B&B companies
        df_platforms['Segment'] = df_platforms.apply(lambda row: find_matching_segment(row, segments_data), axis=1)
        df_platforms['Subcategory'] = df_platforms.apply(lambda row: identify_subcategory(f"{str(row.get('Keywords', '')).lower()} {str(row.get('Description', '')).lower()}", row['Segment']), axis=1)
        print("B&B data classified by Segment and Subcategory.")

    except FileNotFoundError: print(f"Error: B&B file '{bb_file}' not found."); exit()
    except ValueError as e: print(f"Error loading B&B data: {e}"); exit()
    except Exception as e: print(f"An unexpected error occurred loading B&B data: {e}"); exit()


    # 3. Prepare B&B Lookup Dictionary (Revised Structure)
    print("Creating refined lookup dictionary for B&B initiatives...")
    # Structure: subcategory_map[(Segment, Subcategory)] = {investor: {'companies': [list], 'count': N}}
    subcategory_map = {}

    # First pass: Collect all companies per investor per subcategory
    investor_companies_temp = {} # Temp: {(Segment, Subcategory, Investor): {company_set}}
    for _, row in df_platforms.iterrows():
        segment = row['Segment']
        subcategory = row['Subcategory']
        company = row['Companies']
        investors_raw = row.get(investor_col_name, '') # Use found investor column name

        if pd.isna(investors_raw) or investors_raw == '': continue

        investors = [inv.strip() for inv in str(investors_raw).split(',') if inv.strip()]

        for investor in investors:
            key = (segment, subcategory, investor)
            if key not in investor_companies_temp:
                investor_companies_temp[key] = set()
            if pd.notna(company):
                investor_companies_temp[key].add(company)

    # Second pass: Build the final map and count investments
    for (segment, subcategory, investor), companies_set in investor_companies_temp.items():
        full_subcategory_key = (segment, subcategory)
        if full_subcategory_key not in subcategory_map:
            subcategory_map[full_subcategory_key] = {}

        # Get the total count for this investor *in this subcategory*
        investment_count = len(companies_set)

        subcategory_map[full_subcategory_key][investor] = {
            'companies': sorted(list(companies_set)),
            'count': investment_count
        }

    print("Created refined lookup dictionary.")


    # 4. Process Targets: Apply NACE mapping, Filter, Calculate Score, Format Financials
    print("\nProcessing target companies...")
    df_target['Segment'], df_target['Subcategory'] = zip(*df_target['NACE'].apply(get_subcategory_from_nace))

    original_target_count = len(df_target)
    # Create a copy to avoid SettingWithCopyWarning
    df_target_filtered = df_target.dropna(subset=['Segment', 'Subcategory']).copy()

    # Filter out 'General' subcategory BEFORE other calculations
    count_before_general_filter = len(df_target_filtered)
    df_target_filtered = df_target_filtered[df_target_filtered['Subcategory'] != 'General']
    count_after_general_filter = len(df_target_filtered)
    print(f"Filtered targets: {count_after_general_filter} relevant companies found (out of {original_target_count}, removed {count_before_general_filter - count_after_general_filter} 'General' category).")


    # Calculate Score & Format Financials safely using .loc on the filtered copy
    df_target_filtered.loc[:, 'Score_Numeric'] = pd.to_numeric(df_target_filtered['Score'], errors='coerce').fillna(0)
    df_target_filtered.loc[:, 'Exit Probability (%)'] = (df_target_filtered['Score_Numeric'].clip(0, 10) / 10 * 100).round(1)

    # Apply formatting using the revised function
    df_target_filtered.loc[:, 'Revenue (NOK k)'] = df_target_filtered['Revenue'].apply(format_nok_thousands)
    df_target_filtered.loc[:, 'EBIT (NOK k)'] = df_target_filtered['EBIT'].apply(format_nok_thousands)

    # 5. Match Targets to B&B Initiatives (Revised)
    print("Matching targets to potential B&B acquirers...")
    df_target_filtered.loc[:,'Potential Acquirers'] = df_target_filtered.apply(
        lambda row: subcategory_map.get((row['Segment'], row['Subcategory']), {}), axis=1
    )

    # 6. Sort Data
    df_final = df_target_filtered.sort_values('Exit Probability (%)', ascending=False)

    # Get unique subcategories *after* filtering for buttons
    unique_subcategories = sorted(df_final['Subcategory'].unique())

    print("Targets matched and sorted.")

    # 7. Generate HTML Report (Revised)
    print("\nGenerating HTML report 'potential_targets_overview.html'...")
    # --- HTML Structure ---
    html_start = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Potential B&B Targets Overview</title>
    <style>
        :root {{ --primary: #2c3e50; --secondary: #34495e; --accent: #3498db; --light: #ecf0f1; --dark: #2c3e50; --success: #2ecc71; --warning: #f39c12; --danger: #e74c3c; }}
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; background-color: #f4f7f6; color: var(--dark); font-size: 14px; }}
        .container {{ max-width: 1600px; margin: 20px auto; padding: 20px; background-color: #fff; box-shadow: 0 0 15px rgba(0,0,0,0.1); border-radius: 8px; }}
        header {{ background-color: var(--primary); color: white; padding: 1rem 0; margin-bottom: 1.5rem; text-align: center; border-radius: 8px 8px 0 0; }}
        h1 {{ margin: 0; font-size: 1.8rem; }}
        h2 {{ color: var(--primary); border-bottom: 2px solid var(--accent); padding-bottom: 5px; margin-top: 1.5rem; margin-bottom: 1rem; font-size: 1.4rem; }}
        .filter-container {{ margin-bottom: 1.5rem; padding-bottom: 1rem; border-bottom: 1px solid #eee; }}
        .filter-container button {{ background-color: #e9ecef; border: 1px solid #ced4da; color: var(--dark); padding: 6px 12px; margin: 0 4px 4px 0; border-radius: 4px; cursor: pointer; transition: background-color 0.2s; font-size: 0.85rem; }}
        .filter-container button.active {{ background-color: var(--accent); color: white; border-color: var(--accent); font-weight: bold; }}
        .filter-container button:hover {{ background-color: #dee2e6; }}
        .filter-container button.active:hover {{ background-color: #2980b9; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 1rem; }}
        th, td {{ border: 1px solid #ddd; padding: 8px 10px; text-align: left; vertical-align: top; }}
        th {{ background-color: var(--secondary); color: white; white-space: nowrap; }}
        tr.filtered-out {{ display: none; }} /* Class to hide rows */
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        /* tr:hover {{ background-color: #f1f1f1; }} */ /* Can be distracting with filtering */
        .score {{ font-weight: bold; text-align: right; white-space: nowrap; }}
        .score-bar-container {{ width: 80px; height: 12px; background-color: #e0e0e0; border-radius: 3px; overflow: hidden; display: inline-block; vertical-align: middle; margin-left: 5px; }}
        .score-bar {{ height: 100%; background-color: var(--success); border-radius: 3px 0 0 3px; transition: width 0.5s ease-in-out; }}

        /* Acquirer List Styling */
        .acquirer-cell {{ position: relative; }} /* Container for list and toggle button */
        .acquirer-list-wrapper {{ max-height: 8em; /* Approx 6 lines */ overflow: hidden; transition: max-height 0.3s ease-out; }}
        .acquirer-list-wrapper.expanded {{ max-height: 1000px; /* Allow full expansion */ transition: max-height 0.5s ease-in; }}
        .acquirer-list-wrapper ul {{ margin: 0; padding-left: 0; list-style-type: none; }}
        .acquirer-list-wrapper > ul > li {{ margin-bottom: 8px; }} /* Space between investors */
        .acquirer-list-wrapper strong {{ color: var(--primary); }}
        .acquirer-list-wrapper .investor-companies ul {{ margin: 2px 0 0 15px; padding: 0; list-style-type: disc; }}
        .acquirer-list-wrapper .investor-companies li {{ margin-bottom: 2px; font-size: 0.9em; }}
        .toggle-acquirers {{ display: none; /* Hidden by default */ cursor: pointer; color: var(--accent); font-size: 0.85em; margin-top: 5px; text-decoration: underline; }}
        .acquirer-list-wrapper.overflowing + .toggle-acquirers {{ display: block; /* Show only if needed */ }}

        .no-match {{ color: #888; font-style: italic; }}
        .financials {{ text-align: right; white-space: nowrap; }}
        .footer {{ text-align: center; margin-top: 2rem; padding-top: 1rem; border-top: 1px solid #eee; font-size: 0.85rem; color: #777; }}
    </style>
</head>
<body>
    <div class="container">
        <header><h1>Potential Buy & Build Targets</h1></header>

        <h2>Filter by Subcategory</h2>
        <div class="filter-container">
            <button class="filter-button active" data-subcategory="all">Show All</button>
"""
    # Add buttons for each unique subcategory (excluding General)
    for subcat in unique_subcategories:
         html_start += f'            <button class="filter-button" data-subcategory="{subcat}">{subcat}</button>\n'

    html_start += """
        </div>

        <h2>Targets Mapped to B&B Subcategories (Sorted by Exit Probability)</h2>
        <table id="targets-table">
            <thead>
                <tr>
                    <th>#</th>
                    <th>Target Company</th>
                    <th>Segment</th>
                    <th data-column="subcategory">Subcategory</th>
                    <th>Revenue (NOK k)</th>
                    <th>EBIT (NOK k)</th>
                    <th>Exit Probability (%)</th>
                    <th>Potential Acquirers</th>
                </tr>
            </thead>
            <tbody>
    """

    html_rows = ""
    if df_final.empty:
        html_rows = '<tr><td colspan="8" style="text-align:center; padding: 20px;">No relevant targets found matching the criteria (after filtering).</td></tr>'
    else:
        df_final_display = df_final.reset_index(drop=True)
        for i, row in df_final_display.iterrows():
            # Score bar color
            score_val = row['Exit Probability (%)']
            if score_val >= 70: score_color = 'var(--success)'
            elif score_val >= 40: score_color = 'var(--warning)'
            else: score_color = 'var(--danger)'

            # Build Acquirer HTML (with wrapper div)
            acquirers_data = row['Potential Acquirers']
            acquirers_list_html = '<ul class="acquirer-list">'
            if acquirers_data:
                sorted_investors = sorted(acquirers_data.keys())
                for investor in sorted_investors:
                    details = acquirers_data[investor]
                    count = details['count']
                    companies = details['companies']
                    acquirers_list_html += f"""
                        <li>
                            <strong>{investor}</strong> ({count} investment{'s' if count > 1 else ''} in subcategory):
                            <div class="investor-companies">
                                <ul>{''.join(f'<li>{c}</li>' for c in companies)}</ul>
                            </div>
                        </li>"""
                acquirers_list_html += "</ul>"
                # Generate the cell content with wrapper and toggle link
                acquirers_cell_content = f'<div class="acquirer-list-wrapper">{acquirers_list_html}</div><a href="#" class="toggle-acquirers">Show More...</a>'
            else:
                acquirers_cell_content = '<span class="no-match">None found</span>'


            html_rows += f"""
                <tr data-subcategory="{row['Subcategory']}">
                    <td>{i+1}</td>
                    <td>{row['TargetName']}</td>
                    <td>{row['Segment']}</td>
                    <td>{row['Subcategory']}</td>
                    <td class="financials">{row['Revenue (NOK k)']}</td>
                    <td class="financials">{row['EBIT (NOK k)']}</td>
                    <td class="score">
                        {score_val}%
                        <div class="score-bar-container">
                           <div class="score-bar" style="width: {score_val}%; background-color: {score_color};"></div>
                        </div>
                    </td>
                    <td class="acquirer-cell">{acquirers_cell_content}</td>
                </tr>
            """

    # --- HTML End + JavaScript ---
    html_end = """
            </tbody>
        </table>
        <div class="footer">
            Report generated on: """ + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + """ <br>
            Targets from 'main_target_framework.xlsx', B&B data from 'B&B platforms and addons.xlsx'.
        </div>
    </div>

    <script>
        // Subcategory Filtering
        const filterButtons = document.querySelectorAll('.filter-button');
        const tableRows = document.querySelectorAll('#targets-table tbody tr');

        filterButtons.forEach(button => {
            button.addEventListener('click', () => {
                const targetSubcategory = button.getAttribute('data-subcategory');
                filterButtons.forEach(btn => btn.classList.remove('active'));
                button.classList.add('active');
                tableRows.forEach(row => {
                    if (row.hasAttribute('data-subcategory')) {
                        const rowSubcategory = row.getAttribute('data-subcategory');
                        row.classList.toggle('filtered-out', !(targetSubcategory === 'all' || rowSubcategory === targetSubcategory));
                    }
                });
            });
        });

        // Acquirer List Toggling
        document.addEventListener('DOMContentLoaded', () => {
            document.querySelectorAll('.acquirer-list-wrapper').forEach(wrapper => {
                // Check if content overflows the max-height
                if (wrapper.scrollHeight > wrapper.offsetHeight) {
                    wrapper.classList.add('overflowing'); // Mark as overflowing to show button via CSS
                }
            });

            document.querySelectorAll('.toggle-acquirers').forEach(toggleLink => {
                toggleLink.addEventListener('click', (event) => {
                    event.preventDefault(); // Prevent jumping to top
                    const wrapper = toggleLink.previousElementSibling; // The .acquirer-list-wrapper
                    const isExpanded = wrapper.classList.contains('expanded');

                    wrapper.classList.toggle('expanded');
                    toggleLink.textContent = isExpanded ? 'Show More...' : 'Show Less';
                });
            });
        });
    </script>
</body>
</html>
    """

    # Combine and save HTML
    full_html_content = html_start + html_rows + html_end
    output_html_file = 'potential_targets_overview.html'
    try:
        with open(output_html_file, 'w', encoding='utf-8') as f:
            f.write(full_html_content)
        print(f"\nSuccessfully generated HTML report: '{output_html_file}'")
    except Exception as e:
        print(f"\nError writing HTML file: {e}")

    # Optionally save the final mapped data to Excel/CSV
    try:
         output_excel_file = 'mapped_targets_output.xlsx'
         cols_to_output = [
              'TargetName', 'Segment', 'Subcategory',
              'Revenue (NOK k)', 'EBIT (NOK k)', 'Exit Probability (%)',
              'NACE', 'Score', 'Potential Acquirers' # Keep raw data for excel
         ]
         # Use df_final which has the filtered data and raw Potential Acquirers
         df_final[cols_to_output].to_excel(output_excel_file, index=False)
         print(f"Mapped data also saved to '{output_excel_file}'")
    except Exception as e:
         print(f"Could not save Excel output: {e}") 