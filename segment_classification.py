import pandas as pd
import numpy as np
import datetime
import re

# Define the segments and their associated keywords
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

# Add subcategory definitions after the segments_data definition
subcategory_keywords = {
    'Mechanical, Electrical & HVAC': {
        'HVAC': ['hvac', 'ventilation', 'air', 'climate', 'cooling'],
        'Electrical': ['electrical', 'power', 'lighting', 'elektro'],
        'Plumbing': ['plumbing', 'pipe', 'drainage', 'sanitary'],
        'Heating': ['heating', 'geothermal', 'heat pump', 'boiler']
    },
    'Core Construction & Civil Engineering': {
        'Civil Engineering': ['civil engineering', 'infrastructure', 'road', 'tunnel', 'excavation', 'groundwork'],
        'Building Construction': ['building', 'construction company', 'house builder', 'residential'],
        'Property Services': ['property', 'real estate', 'renovation', 'maintenance']
    },
    'Specialized Trades': {
        'Flooring': ['floor', 'tiling', 'tiles'],
        'Carpentry': ['carpentry', 'wood', 'timber'],
        'Roofing': ['roof', 'roofing'],
        'Painting': ['paint', 'coating', 'surface']
    },
    'Industrial Services & Manufacturing Support': {
        'Welding & Metalwork': ['welding', 'steel', 'metal', 'forging'],
        'Concrete': ['concrete', 'cement'],
        'Maintenance': ['maintenance', 'repair', 'service']
    }
}

# Add a default subcategory for all other segments
for segment in segments_data:
    if segment not in subcategory_keywords:
        subcategory_keywords[segment] = {'General': [segment.lower()]}

# Function to identify subcategory based on text
def identify_subcategory(text, segment):
    # Default subcategory is just the segment name
    default_subcategory = segment.split(' ')[0]
    
    if not text or not isinstance(text, str):
        return default_subcategory
    
    text = text.lower()
    
    # Get subcategory definitions for this segment
    if segment in subcategory_keywords:
        subcategories = subcategory_keywords[segment]
        for subcategory, keywords in subcategories.items():
            for keyword in keywords:
                if keyword.lower() in text:
                    return subcategory
    
    # If no subcategory found, return default
    return default_subcategory

# Create a DataFrame for segments
df_segments = pd.DataFrame([(segment, keyword) 
                          for segment, keywords in segments_data.items() 
                          for keyword in keywords],
                         columns=['Segment', 'Keyword'])

# Read the Excel files, using row 7 as headers
print("\nReading Excel files...")
df_platforms = pd.read_excel('B&B platforms and addons.xlsx', header=6, engine='openpyxl', sheet_name='Data')
df_investors = pd.read_excel('B&B investors.xlsx', header=6, engine='openpyxl')

# Clean up the data by removing empty rows and columns
df_platforms = df_platforms.dropna(how='all').dropna(axis=1, how='all')
df_investors = df_investors.dropna(how='all').dropna(axis=1, how='all')

# --- Debug: Print platform columns to confirm new column name ---
print("\nDetected columns in Platforms/Addons sheet:")
print(df_platforms.columns.tolist())
# --- End Debug ---

print("\nPlatforms and Add-ons DataFrame structure:")
print("\nColumns:")
for col in df_platforms.columns:
    print(f"- {col}")
print("\nFirst 3 rows of data:")
print(df_platforms.head(3))

print("\n" + "="*80 + "\n")

print("Investors DataFrame structure:")
print("\nColumns:")
for col in df_investors.columns:
    print(f"- {col}")
print("\nFirst 3 rows of data:")
print(df_investors.head(3))

# Display the number of records in each DataFrame
print(f"\nNumber of platform/addon records: {len(df_platforms)}")
print(f"Number of investor records: {len(df_investors)}")

# Function to find matching segment based on keywords and description
def find_matching_segment(row, segments_data):
    text_to_search = f"{str(row['Keywords']).lower()} {str(row['Description']).lower()}"
    matching_segments = []
    
    for segment, keywords in segments_data.items():
        for keyword in keywords:
            if keyword.lower() in text_to_search:
                matching_segments.append(segment)
                break
    
    if matching_segments:
        return matching_segments[0]  # Return the first matching segment
    return "Other"

# Add segment classification to platforms DataFrame
print("\nClassifying companies into segments...")
df_platforms['Segment'] = df_platforms.apply(lambda row: find_matching_segment(row, segments_data), axis=1)

# Create B&B initiatives overview
print("\nCreating B&B initiatives overview...")
initiatives = []

# Define the expected date column name (WILL BE CONFIRMED/UPDATED based on printout below)
DATE_COLUMN = 'Last Financing Date'

# --- Confirmation step: Check if the identified column K header matches DATE_COLUMN ---
actual_columns = df_platforms.columns.tolist()
if len(actual_columns) > 10: # Check if column K (index 10) exists
    ACTUAL_DATE_COLUMN_NAME = actual_columns[10]
    print(f"\nDEBUG: Identified Header for Column K (index 10): '{ACTUAL_DATE_COLUMN_NAME}'")
    if ACTUAL_DATE_COLUMN_NAME != DATE_COLUMN:
        print(f"INFO: Updating DATE_COLUMN to '{ACTUAL_DATE_COLUMN_NAME}' based on file header.")
        DATE_COLUMN = ACTUAL_DATE_COLUMN_NAME # Update the variable
    else:
        print(f"INFO: Confirmed DATE_COLUMN '{DATE_COLUMN}' matches header in Column K.")
else:
    print("\nWARNING: Could not identify column K (index 10) header.")
# --- End Confirmation Step ---


# Check if the date column exists in the DataFrame columns
has_date_column = DATE_COLUMN in df_platforms.columns
if not has_date_column:
    print(f"\nWarning: Column '{DATE_COLUMN}' not found after header check. Company years will not be added.")

# Group companies by investor and segment, handling multiple investors per company
for _, row in df_platforms.iterrows():
    if pd.isna(row['All investors']):
        continue
    
    # Split investors and clean up any whitespace
    investors = [inv.strip() for inv in str(row['All investors']).split(',')] # Added str() for safety
    
    # Format company name with year
    company_name = row['Companies']
    if has_date_column and pd.notna(row[DATE_COLUMN]):
        try:
            # Convert to datetime and extract year
            year = pd.to_datetime(row[DATE_COLUMN]).year
            company_name = f"{company_name} ({year})"
        except (ValueError, TypeError) as e:
            # Handle cases where conversion fails
            print(f"Warning: Could not parse year from date '{row[DATE_COLUMN]}' for company '{row['Companies']}'. Error: {e}")
    
    # Create an entry for each individual investor
    for investor in investors:
        initiatives.append({
            'Investor': investor,
            'Segment': row['Segment'],
            'Company': company_name
        })

# Create initiatives DataFrame
df_initiatives = pd.DataFrame(initiatives)

# Create summary by grouping by Investor and Segment
df_summary = df_initiatives.groupby(['Investor', 'Segment']).agg({
    'Company': lambda x: list(x)
}).reset_index()

# Add number of companies column
df_summary['Number of Companies'] = df_summary['Company'].apply(len)
# Convert company list to comma-separated string
df_summary['Companies'] = df_summary['Company'].apply(lambda x: ', '.join(x))
# Drop the list column
df_summary = df_summary.drop('Company', axis=1)

# Sort by Investor and Number of Companies
df_summary = df_summary.sort_values(['Investor', 'Number of Companies'], ascending=[True, False])

# Save results to Excel
print("\nSaving results...")
with pd.ExcelWriter('bb_initiatives_overview.xlsx') as writer:
    # Save company classifications (include the date column)
    columns_to_save = ['Companies', 'All investors', 'Segment', 'Description']
    if has_date_column:
        columns_to_save.append(DATE_COLUMN)
    else:
        print(f"Warning: Date column '{DATE_COLUMN}' not found, will not be included in 'Company Classifications' sheet.")
        
    df_platforms[columns_to_save].to_excel(
        writer, sheet_name='Company Classifications', index=False
    )
    
    # Save B&B initiatives overview (with Company Name (Year))
    df_summary.to_excel(writer, sheet_name='B&B Initiatives', index=False)
    
    # Save segment statistics
    segment_stats = df_platforms['Segment'].value_counts().reset_index()
    segment_stats.columns = ['Segment', 'Number of Companies']
    segment_stats.to_excel(writer, sheet_name='Segment Statistics', index=False)
    
    # Save investor statistics
    investor_stats = df_summary.groupby('Investor')['Number of Companies'].sum().sort_values(ascending=False).reset_index()
    investor_stats.columns = ['Investor', 'Total Companies']
    investor_stats.to_excel(writer, sheet_name='Investor Statistics', index=False)

print("\nResults have been saved to 'bb_initiatives_overview.xlsx'")

# Display some summary statistics
print("\nSegment Distribution:")
print(df_platforms['Segment'].value_counts())

print("\nTop 5 Investors by number of companies:")
print(df_summary.groupby('Investor')['Number of Companies'].sum().sort_values(ascending=False).head())

# Save the segments DataFrame to CSV for future use
df_segments.to_csv('construction_segments.csv', index=False)
print("\nDataFrame has been saved to 'construction_segments.csv'")

# ------------------------------------------------
# HTML Report Generation
# ------------------------------------------------

print("\nGenerating HTML report...")

# Modify the company_details dictionary to include subcategory
company_details = {}
for _, row in df_platforms.iterrows():
    company_name = row['Companies']
    company_id = row['Company ID']
    segment = row['Segment']
    description = str(row['Description']) if pd.notna(row['Description']) else ''
    keywords = str(row['Keywords']) if pd.notna(row['Keywords']) else ''
    
    # Determine subcategory using both description and keywords
    text_for_subcategory = f"{description} {keywords}"
    subcategory = identify_subcategory(text_for_subcategory, segment)
    
    # Get Revenue and EBITDA, handle potential missing columns or NaN values
    # and format if numeric
    revenue_val = row['Revenue'] if 'Revenue' in row and pd.notna(row['Revenue']) else None
    ebitda_val = row['EBITDA'] if 'EBITDA' in row and pd.notna(row['EBITDA']) else None
    
    revenue_str = f"€{revenue_val}m" if revenue_val is not None else 'N/A'
    ebitda_str = f"€{ebitda_val}m" if ebitda_val is not None else 'N/A'
    
    # Attempt to handle cases where the value might already contain 'm' or other chars
    # This is a basic check, might need refinement based on actual data patterns
    if isinstance(revenue_val, str) and not revenue_val.replace('.','',1).isdigit():
        revenue_str = str(revenue_val) # Keep original string if it's not purely numeric
    elif revenue_val is not None:
         revenue_str = f"€{revenue_val}m" 
    else:
        revenue_str = 'N/A'
        
    if isinstance(ebitda_val, str) and not ebitda_val.replace('.','',1).isdigit():
        ebitda_str = str(ebitda_val) # Keep original string if it's not purely numeric
    elif ebitda_val is not None:
         ebitda_str = f"€{ebitda_val}m" 
    else:
        ebitda_str = 'N/A'

    company_details[company_name] = {
        'id': company_id,  # <-- Store Company ID here
        'description': description,
        'date': pd.to_datetime(row[DATE_COLUMN]).strftime('%Y-%m-%d') if pd.notna(row[DATE_COLUMN]) else 'Unknown',
        'year': pd.to_datetime(row[DATE_COLUMN]).year if pd.notna(row[DATE_COLUMN]) else 'Unknown',
        'subcategory': subcategory,
        'revenue': revenue_str,  # Use formatted revenue string
        'ebitda': ebitda_str    # Use formatted EBITDA string
    }

# Prepare data for HTML: Group by segment, then by investor
html_data = {}
for segment in df_platforms['Segment'].unique():
    html_data[segment] = {}
    
    # Get all initiatives for this segment
    segment_initiatives = df_summary[df_summary['Segment'] == segment]
    
    for _, row in segment_initiatives.iterrows():
        investor = row['Investor']
        companies_str = row['Companies']
        companies = [c.strip() for c in companies_str.split(',')]
        
        # Extract company details
        company_list = []
        for company in companies:
            # Extract base company name (remove the year part if present)
            base_name = company.split(' (')[0] if ' (' in company else company
            
            if base_name in company_details:
                retrieved_id = company_details[base_name].get('id', 'MISSING_ID') 
                company_list.append({
                    'id': retrieved_id, 
                    'base_name': base_name,
                    'year': company_details[base_name].get('year', 'Unknown'),
                    'date': company_details[base_name].get('date', 'Unknown'),
                    'description': company_details[base_name].get('description', 'N/A'),
                    'subcategory': company_details[base_name].get('subcategory', 'Unknown'),
                    'revenue': company_details[base_name].get('revenue', 'N/A'), # Add revenue
                    'ebitda': company_details[base_name].get('ebitda', 'N/A')   # Add EBITDA
                })
            else:
                # Fallback logic
                fallback_id = company_details.get(company, {}).get('id', 'MISSING_ID')
                company_list.append({
                    'id': fallback_id, 
                    'base_name': base_name,
                    'year': company_details.get(company, {}).get('year', 'Unknown'),
                    'date': company_details.get(company, {}).get('date', 'Unknown'),
                    'description': company_details.get(company, {}).get('description', 'N/A'),
                    'subcategory': company_details.get(company, {}).get('subcategory', 'Unknown'),
                    'revenue': company_details.get(company, {}).get('revenue', 'N/A'), # Add revenue (fallback)
                    'ebitda': company_details.get(company, {}).get('ebitda', 'N/A')   # Add EBITDA (fallback)
                })
        
        # Add the company list to the nested structure
        html_data[segment][investor] = company_list

# Order segments by number of companies
segments_ordered = df_platforms['Segment'].value_counts().index.tolist()

# Calculate some statistics for the dashboard
total_initiatives = len(df_summary)
total_companies = len(df_platforms)
total_investors = len(df_initiatives['Investor'].unique())
total_segments = len(df_platforms['Segment'].unique())
last_updated = datetime.datetime.now().strftime('%Y-%m-%d')

# Define color palette for segments
segment_colors = {
    'Core Construction & Civil Engineering': '#3498db',  # Blue
    'Specialized Trades': '#e74c3c',  # Red
    'Mechanical, Electrical & HVAC': '#2ecc71',  # Green
    'Marine, Offshore & Energy': '#f39c12',  # Orange
    'Industrial Services & Manufacturing Support': '#9b59b6',  # Purple
    'Building Products & Materials': '#1abc9c',  # Turquoise
    'Tech & Software for Construction': '#34495e',  # Dark Blue
    'Consulting, Advisory & Project Management': '#e67e22',  # Dark Orange
    'Equipment Rental & Heavy Machinery': '#27ae60',  # Dark Green
    'Facility Services & Real Estate Ops': '#c0392b',  # Dark Red
    'Safety & Monitoring Systems': '#8e44ad',  # Dark Purple
    'Environmental & Waste Management': '#16a085',  # Dark Turquoise
    'Infrastructure & Public Works': '#f1c40f',  # Yellow
    'Interior Design & Furnishing': '#7f8c8d',  # Gray
    'Other': '#95a5a6'  # Light Gray
}

# After creating html_data but before HTML generation, add platform/add-on identification logic
# First, normalize the revenue and EBITDA columns
# This block should go before the HTML template creation code begins

print("Identifying platform vs. add-on investments...")

# Extract numeric values from Revenue and EBITDA columns (if available)
def extract_numeric(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        # Remove currency symbols, commas, 'm', 'k', etc. and try to convert to float
        clean_value = re.sub(r'[^\d.]', '', value.replace('k', '').replace('m', ''))
        try:
            return float(clean_value)
        except:
            return np.nan
    return np.nan

# If the platforms DataFrame has Revenue and EBITDA columns, use them for classification
revenue_col = None
ebitda_col = None

for col in df_platforms.columns:
    if 'revenue' in col.lower():
        revenue_col = col
    elif 'ebitda' in col.lower():
        ebitda_col = col

# Criteria for platform identification
# 1. First acquisition in a segment by an investor
# 2. Larger revenue/EBITDA compared to other acquisitions (if available)
# 3. Earlier acquisition date

# Create a dictionary to store platform status of each company
platform_status = {}

# Track the first company acquired by each investor in each segment
first_acquisitions = {}

# Process companies by acquisition date
if 'Last financing date' in df_platforms.columns:
    print("Using 'Last financing date' for platform/add-on classification")
    # Convert date to datetime for sorting
    df_platforms['Last financing date'] = pd.to_datetime(df_platforms['Last financing date'], errors='coerce')
    
    # Sort by date (earliest first)
    sorted_companies = df_platforms.sort_values('Last financing date')
    
    # Extract numeric revenue/EBITDA if available (we'll use this for ranking in case of missing dates)
    if revenue_col:
        sorted_companies['Revenue_Numeric'] = sorted_companies[revenue_col].apply(extract_numeric)
    
    if ebitda_col:
        sorted_companies['EBITDA_Numeric'] = sorted_companies[ebitda_col].apply(extract_numeric)
    
    # Create a dictionary to track first acquisition by each investor in each segment
    investor_first_acquisition = {}
    
    # First pass: identify the first acquisition in each segment by each investor
    for _, company in sorted_companies.iterrows():
        company_id = company['Company ID']
        company_name = company['Companies']
        segment = company['Segment']
        investor_list = str(company['All investors']).split(', ')
        
        acquisition_date = company['Last financing date']
        if pd.isna(acquisition_date):
            continue  # Skip companies with no acquisition date
            
        # Generate a unique key for each company
        key = f"{company_id}_{company_name}"
        
        # Default all companies to add-ons initially
        platform_status[key] = 'Add-on'
        
        # Check for each investor
        for investor in investor_list:
            investor = investor.strip()
            if not investor or investor == 'nan':
                continue
                
            investor_segment_key = f"{investor}_{segment}"
            
            # If this is the first time we've seen this investor-segment combination,
            # or if this acquisition is earlier than the previously recorded one
            if (investor_segment_key not in investor_first_acquisition or 
                acquisition_date < investor_first_acquisition[investor_segment_key][1]):
                investor_first_acquisition[investor_segment_key] = (key, acquisition_date)
    
    # Second pass: mark the earliest acquisition in each segment by each investor as a platform
    for investor_segment_key, (company_key, _) in investor_first_acquisition.items():
        platform_status[company_key] = 'Platform'
        
    # Special case: if there's a particularly large company (5x+ average revenue in segment),
    # also mark it as a platform even if it's not the first acquisition
    if revenue_col:
        # Calculate average revenue by segment
        segment_avg_revenue = {}
        for segment in sorted_companies['Segment'].unique():
            segment_companies = sorted_companies[sorted_companies['Segment'] == segment]
            if 'Revenue_Numeric' in segment_companies.columns:
                revenues = segment_companies['Revenue_Numeric'].dropna()
                if len(revenues) > 0:
                    segment_avg_revenue[segment] = revenues.mean()
        
        # Identify companies with significantly higher revenue than segment average
        for _, company in sorted_companies.iterrows():
            if 'Revenue_Numeric' in sorted_companies.columns:
                segment = company['Segment']
                revenue = company['Revenue_Numeric']
                company_id = company['Company ID']
                company_name = company['Companies']
                key = f"{company_id}_{company_name}"
                
                if (segment in segment_avg_revenue and 
                    not pd.isna(revenue) and 
                    revenue > segment_avg_revenue[segment] * 5):
                    platform_status[key] = 'Platform'  # Mark as platform if revenue is 5x+ segment average
else:
    print("Warning: 'Last Financing Date' column not found. Using basic classification.")
    # Use a simpler approach without date ordering
    for segment in html_data:
        for investor, companies in html_data[segment].items():
            if len(companies) > 0:
                # Check what fields are available in the company data
                sample_company = companies[0]
                # Find the company identifier field
                company_id_field = None
                if 'company_id' in sample_company:
                    company_id_field = 'company_id'
                elif 'id' in sample_company:
                    company_id_field = 'id'
                
                # Mark the first company as platform
                if company_id_field:
                    first_key = f"{sample_company[company_id_field]}_{sample_company['base_name']}"
                else:
                    # If no ID field, just use the name
                    first_key = sample_company['base_name']
                platform_status[first_key] = 'Platform'
                
                # Mark the rest as add-ons
                for i in range(1, len(companies)):
                    company = companies[i]
                    if company_id_field:
                        company_key = f"{company[company_id_field]}_{company['base_name']}"
                    else:
                        company_key = company['base_name']
                    platform_status[company_key] = 'Add-on'

# Update the html_data with investment_type
print(f"Identified {list(platform_status.values()).count('Platform')} platforms and {list(platform_status.values()).count('Add-on')} add-ons")

for segment in html_data:
    for investor, companies in html_data[segment].items():
        for company in companies:
            # Find the company identifier field
            company_id_field = None
            if 'company_id' in company:
                company_id_field = 'company_id'
            elif 'id' in company:
                company_id_field = 'id'
            
            # Generate key for lookup
            if company_id_field:
                key = f"{company[company_id_field]}_{company['base_name']}"
            else:
                key = company['base_name']
                
            company['investment_type'] = platform_status.get(key, 'Unknown')

# Modify how company cards are generated in the HTML template
segments_html = ""

# Add segments to HTML
for segment in segments_ordered:
    if segment not in html_data:
        continue
        
    segment_companies_count = sum(len(companies) for companies in html_data[segment].values())
    segment_color = segment_colors.get(segment, '#3498db')  # Default to blue if segment not in colors dict
    
    segments_html += f'''
            <div class="segment-section" data-segment="{segment}" style="border-top-color: {segment_color}">
                <div class="segment-header">
                    <div class="segment-name">{segment}</div>
                    <div class="segment-count" style="background-color: {segment_color}">{segment_companies_count} companies</div>
                </div>
                <div class="subcategory-filters" data-segment="{segment}">
                    <button class="subcategory-filter active" data-subcategory="all" style="border-color: {segment_color}">All</button>
                </div>
    '''
    
    # Order investors by number of companies (descending)
    sorted_investors = sorted(html_data[segment].items(), 
                             key=lambda x: len(x[1]), 
                             reverse=True)
    
    # Determine if we need a "Show More" button
    has_more_than_five = len(sorted_investors) > 5
    
    for idx, (investor, companies) in enumerate(sorted_investors):
        # Add class to hide investors beyond the first 5
        extra_class = ' hidden-investor' if idx >= 5 else ''
        
        # --- Sort companies by year (earliest first) ---
        def sort_key(company):
            year = company.get('year')
            # Treat 'Unknown' or non-integer years as very large numbers for sorting
            if isinstance(year, int):
                return year
            return float('-inf') # Use negative infinity for unknowns to place them last in descending sort
            
        sorted_companies_by_year = sorted(companies, key=sort_key, reverse=True) # Add reverse=True
        # --- End sorting ---
        
        segments_html += f'''
                <div class="investor-section{extra_class}" data-investor="{investor}">
                    <div class="investor-name">{investor}</div>
                    <div class="company-list">
        '''
        
        # Use the sorted list here
        for company in sorted_companies_by_year:
            # Add platform/add-on badge
            investment_type = company.get('investment_type', 'Unknown')
            badge_class = 'platform-badge' if investment_type == 'Platform' else 'addon-badge'
            # Ensure badge only shows if type is known (Platform or Add-on)
            badge_html = ''
            if investment_type in ['Platform', 'Add-on']:
                 badge_html = f'<div class="investment-type-badge {badge_class}">{investment_type}</div>'
            
            segments_html += f'''
                        <div class="company-card" data-company="{company['base_name']}" data-investment-type="{investment_type}">
                            <div class="acquisition-year" style="background-color: {segment_color}">{company['year']}</div>
                            {badge_html}
                            <div class="subcategory-tag" style="background-color: {segment_color}">{company['subcategory']}</div>
                            <div class="company-name">{company['base_name']}</div>
                            <button class="show-description" style="background-color: {segment_color}">Show Details</button>
                            <div class="company-description">
                                <p><strong>Acquired:</strong> {company['date']}</p>
                                <p><strong>Revenue:</strong> {company.get('revenue', 'N/A')}</p>
                                <p><strong>EBITDA:</strong> {company.get('ebitda', 'N/A')}</p>
                                <p><strong>Description:</strong> {company['description']}</p>
                            </div>
                        </div>
            '''
        
        segments_html += '''
                    </div>
                </div>
        '''
    
    # Add "Show More" button if there are more than 5 investors
    if has_more_than_five:
        segments_html += f'''
                <button class="show-more-button" data-segment="{segment}" data-expanded="false">
                    Show More Investors ({len(sorted_investors) - 5} more)
                </button>
        '''
    
    segments_html += '''
            </div>
    '''

# Update the CSS to include platform/add-on badge styles
css_with_badges = '''
        .investment-type-badge {
            /* Remove absolute positioning */
            /* position: absolute; */ 
            /* top: 0.5rem; */
            /* left: 0.5rem; */
            display: inline-block; /* Make it flow with text */
            margin-right: 0.5rem; /* Add some space */
            margin-bottom: 0.3rem; /* Add space below */
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            font-size: 0.7rem;
            font-weight: bold;
            color: white;
            text-transform: uppercase;
            vertical-align: middle; /* Align with other inline elements */
        }
        
        .platform-badge {
            background-color: #2ecc71;  /* Green */
        }
        
        .addon-badge {
            background-color: #e74c3c;  /* Red */
        }
'''

# Create the HTML template with updated CSS and content
html_template = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Buy & Build Initiatives Overview - Construction & Engineering</title>
    <style>
        :root {{
            --primary: #2c3e50;
            --secondary: #34495e;
            --accent: #3498db;
            --light: #ecf0f1;
            --dark: #2c3e50;
            --success: #2ecc71;
            --info: #3498db;
            --warning: #f39c12;
            --danger: #e74c3c;
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}
        
        body {{
            background-color: #f5f7fa;
            color: var(--dark);
            line-height: 1.6;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        header {{
            background-color: var(--primary);
            color: white;
            padding: 1rem 0;
            margin-bottom: 2rem;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }}
        
        h1 {{
            text-align: center;
            font-size: 2.5rem;
            margin-bottom: 0.5rem;
        }}
        
        h2 {{
            color: var(--primary);
            margin: 1.5rem 0 1rem;
            border-bottom: 2px solid var(--accent);
            padding-bottom: 0.5rem;
        }}
        
        .dashboard {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 2rem;
            flex-wrap: wrap;
        }}
        
        .stat-card {{
            background-color: white;
            border-radius: 8px;
            padding: 1.5rem;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            flex: 1;
            min-width: 200px;
            margin: 0.5rem;
            text-align: center;
            transition: transform 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
        }}
        
        .stat-value {{
            font-size: 2.5rem;
            font-weight: bold;
            color: var(--accent);
            margin: 0.5rem 0;
        }}
        
        .stat-label {{
            font-size: 1rem;
            color: var(--secondary);
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .segment-section {{
            background-color: white;
            border-radius: 8px;
            padding: 1.5rem;
            margin-bottom: 2rem;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            border-top: 5px solid #3498db; /* Default color, will be overridden */
        }}
        
        .segment-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
            cursor: pointer;
        }}
        
        .segment-name {{
            font-size: 1.5rem;
            color: var(--primary);
            font-weight: bold;
        }}
        
        .segment-count {{
            background-color: #3498db; /* Default color, will be overridden */
            color: white;
            padding: 0.25rem 0.75rem;
            border-radius: 50px;
            font-size: 0.9rem;
        }}
        
        .investor-section {{
            margin: 1rem 0 1rem 1.5rem;
            padding-left: 1rem;
            border-left: 3px solid #e1e4e8;
        }}
        
        .investor-section.hidden-investor {{
            display: none;
        }}
        
        .investor-name {{
            font-size: 1.2rem;
            font-weight: 600;
            color: var(--secondary);
            margin-bottom: 0.5rem;
        }}
        
        .company-list {{
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            margin-top: 1rem;
        }}
        
        .company-card {{
            background-color: var(--light);
            border-radius: 6px;
            padding: 1rem;
            flex: 1;
            min-width: 250px;
            position: relative;
            transition: all 0.3s ease;
        }}
        
        .company-card:hover {{
            transform: translateY(-3px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }}
        
        {css_with_badges}
        
        .subcategory-tag {{
            display: inline-block;
            padding: 0.15rem 0.5rem;
            font-size: 0.75rem;
            border-radius: 3px;
            margin-bottom: 0.5rem;
            color: white;
            font-weight: 500;
            opacity: 0.8;
        }}
        
        .company-name {{
            font-weight: bold;
            font-size: 1.1rem;
            margin-bottom: 0.5rem;
            color: var(--dark);
        }}
        
        .acquisition-year {{
            position: absolute;
            /* top: 0.5rem; */ /* Remove top positioning */
            bottom: 0.5rem; /* Add bottom positioning */
            right: 0.5rem;
            color: white;
            padding: 0.2rem 0.5rem;
            border-radius: 4px;
            font-size: 0.8rem;
        }}
        
        .company-description {{
            font-size: 0.9rem;
            color: var(--secondary);
            margin-top: 0.5rem;
            display: none;
        }}
        
        .show-description {{
            background-color: var(--accent);
            color: white;
            border: none;
            padding: 0.4rem 0.8rem;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.8rem;
            margin-top: 0.5rem;
        }}
        
        .show-description:hover {{
            background-color: var(--info);
        }}
        
        .show-more-button {{
            display: block;
            margin: 1rem auto;
            background-color: #f5f7fa;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 0.5rem 1rem;
            cursor: pointer;
            transition: all 0.3s ease;
            font-size: 0.9rem;
        }}
        
        .show-more-button:hover {{
            background-color: #e8eaed;
        }}
        
        .subcategory-filters {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
            margin-bottom: 1rem;
            margin-top: 0.5rem;
            padding-bottom: 0.5rem;
            border-bottom: 1px solid #eee;
        }}
        
        .subcategory-filter {{
            background-color: white;
            color: inherit;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 0.2rem 0.6rem;
            cursor: pointer;
            transition: all 0.3s ease;
            font-size: 0.8rem;
        }}
        
        .subcategory-filter.active {{
            background-color: #f5f7fa;
            font-weight: bold;
            box-shadow: inset 0 0 0 1px currentColor;
        }}
        
        .subcategory-filter:hover {{
            background-color: #f5f7fa;
        }}
        
        .segment-buttons-container {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
            margin-bottom: 2rem;
        }}
        
        .segment-button {{
            background-color: white;
            color: white;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 0.5rem 1rem;
            cursor: pointer;
            transition: all 0.3s ease;
            font-size: 0.9rem;
        }}
        
        .segment-button[data-segment="all"] {{
            background-color: var(--primary);
            border-color: var(--primary);
        }}
        
        .segment-button:hover {{
            opacity: 0.9;
            transform: translateY(-2px);
        }}
        
        .segment-button.active {{
            box-shadow: 0 0 0 2px white, 0 0 0 4px currentColor;
        }}
        
        footer {{
            text-align: center;
            margin-top: 3rem;
            padding: 1rem;
            background-color: var(--primary);
            color: white;
        }}
        
        .hidden {{
            display: none;
        }}
        
        @media (max-width: 768px) {{
            .dashboard {{
                flex-direction: column;
            }}
            
            .stat-card {{
                margin-bottom: 1rem;
            }}
            
            .company-card {{
                min-width: 100%;
            }}
        }}
    </style>
</head>
<body>
    <header>
        <div class="container">
            <h1>Buy & Build Initiatives Overview - Construction & Engineering</h1>
            <p style="text-align: center; color: #ccc;">Analysis of construction industry consolidation in Norway and Sweden</p>
        </div>
    </header>
    
    <div class="container">
        <div class="dashboard">
            <div class="stat-card">
                <div class="stat-label">Total Companies</div>
                <div class="stat-value">{total_companies}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Investors</div>
                <div class="stat-value">{total_investors}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Segments</div>
                <div class="stat-value">{total_segments}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Initiatives</div>
                <div class="stat-value">{total_initiatives}</div>
            </div>
        </div>
        
        <div class="segment-buttons-container">
            <button class="segment-button active" data-segment="all">All Segments</button>
'''

# Add segment buttons (after the "All Segments" button)
for segment in segments_ordered:
    if segment not in html_data:
        continue
    segment_color = segment_colors.get(segment, '#3498db')
    html_template += f'''
            <button class="segment-button" 
                    data-segment="{segment}" 
                    style="background-color: {segment_color}; border-color: {segment_color};">
                {segment}
            </button>'''

html_template += '''
        </div>
        
        <div id="initiatives-container">
'''

# Add the segments HTML
html_template += segments_html

# Close the HTML template
html_template += f'''
        </div>
    </div>
    
    <footer>
        <div class="container">
            <p>Last updated: {last_updated}</p>
            <p>Generated from Buy & Build Initiatives Analysis</p>
        </div>
    </footer>
    
    <script>
        // Toggle description visibility
        document.querySelectorAll('.show-description').forEach(button => {{
            button.addEventListener('click', function() {{
                const description = this.nextElementSibling;
                if (description.style.display === 'block') {{
                    description.style.display = 'none';
                    this.textContent = 'Show Details';
                }} else {{
                    description.style.display = 'block';
                    this.textContent = 'Hide Details';
                }}
            }});
        }});
        
        // Toggle more investors visibility
        document.querySelectorAll('.show-more-button').forEach(button => {{
            button.addEventListener('click', function() {{
                const segmentId = this.getAttribute('data-segment');
                const isExpanded = this.getAttribute('data-expanded') === 'true';
                const segment = this.closest('.segment-section');
                
                // Find hidden investors in this segment
                const hiddenInvestors = segment.querySelectorAll('.hidden-investor');
                
                if (isExpanded) {{
                    // Hide investors
                    hiddenInvestors.forEach(investor => {{
                        investor.classList.add('hidden-investor');
                    }});
                    this.setAttribute('data-expanded', 'false');
                    this.textContent = `Show More Investors (${{hiddenInvestors.length}} more)`;
                }} else {{
                    // Show investors
                    hiddenInvestors.forEach(investor => {{
                        investor.classList.remove('hidden-investor');
                    }});
                    this.setAttribute('data-expanded', 'true');
                    this.textContent = 'Show Less';
                }}
            }});
        }});
        
        // Segment toggle functionality
        const segmentButtons = document.querySelectorAll('.segment-button');
        let activeSegment = 'all';
        
        segmentButtons.forEach(button => {{
            button.addEventListener('click', function() {{
                const segment = this.getAttribute('data-segment');
                
                // If clicking the already active segment, reset to show all
                if (segment === activeSegment && segment !== 'all') {{
                    // Reset - show all segments
                    document.querySelector('button[data-segment="all"]').classList.add('active');
                    this.classList.remove('active');
                    activeSegment = 'all';
                    
                    // Show all segments
                    document.querySelectorAll('.segment-section').forEach(section => {{
                        section.style.display = '';
                    }});
                }} else {{
                    // Update active state
                    segmentButtons.forEach(btn => btn.classList.remove('active'));
                    this.classList.add('active');
                    activeSegment = segment;
                    
                    // Show/hide segments based on selection
                    document.querySelectorAll('.segment-section').forEach(section => {{
                        if (segment === 'all') {{
                            section.style.display = '';
                        }} else {{
                            section.style.display = section.getAttribute('data-segment') === segment ? '' : 'none';
                        }}
                    }});
                }}
            }});
        }});
        
        // Collect all subcategories for each segment
        document.querySelectorAll('.segment-section').forEach(segment => {{
            const segmentId = segment.getAttribute('data-segment');
            const subcategories = new Set();
            
            // Find all subcategories in this segment
            segment.querySelectorAll('.subcategory-tag').forEach(tag => {{
                subcategories.add(tag.textContent);
            }});
            
            // Create filter buttons for each subcategory
            const filterContainer = segment.querySelector('.subcategory-filters');
            if (filterContainer) {{
                subcategories.forEach(subcategory => {{
                    const button = document.createElement('button');
                    button.className = 'subcategory-filter';
                    button.setAttribute('data-subcategory', subcategory);
                    button.textContent = subcategory;
                    button.style.borderColor = getComputedStyle(segment).borderTopColor;
                    filterContainer.appendChild(button);
                }});
                
                // Add Investment Type filters
                const divider = document.createElement('div');
                divider.style.borderTop = '1px solid #eee';
                divider.style.margin = '0.5rem 0';
                divider.style.width = '100%';
                filterContainer.appendChild(divider);
                
                // Add filter heading
                const filterHeading = document.createElement('div');
                filterHeading.textContent = 'Filter by investment type:';
                filterHeading.style.fontSize = '0.8rem';
                filterHeading.style.marginBottom = '0.3rem';
                filterContainer.appendChild(filterHeading);
                
                // Create filter buttons container
                const typeFilters = document.createElement('div');
                typeFilters.style.display = 'flex';
                typeFilters.style.gap = '0.5rem';
                
                // All filter
                const allTypeFilter = document.createElement('button');
                allTypeFilter.className = 'subcategory-filter active';
                allTypeFilter.setAttribute('data-investment-type', 'all');
                allTypeFilter.textContent = 'All Types';
                typeFilters.appendChild(allTypeFilter);
                
                // Platform filter
                const platformFilter = document.createElement('button');
                platformFilter.className = 'subcategory-filter';
                platformFilter.setAttribute('data-investment-type', 'Platform');
                platformFilter.textContent = 'Platforms';
                platformFilter.style.borderColor = '#2ecc71';
                typeFilters.appendChild(platformFilter);
                
                // Add-on filter
                const addonFilter = document.createElement('button');
                addonFilter.className = 'subcategory-filter';
                addonFilter.setAttribute('data-investment-type', 'Add-on');
                addonFilter.textContent = 'Add-ons';
                addonFilter.style.borderColor = '#e74c3c';
                typeFilters.appendChild(addonFilter);
                
                filterContainer.appendChild(typeFilters);
                
                // Add click event listeners to subcategory filters
                segment.querySelectorAll('.subcategory-filter').forEach(button => {{
                    button.addEventListener('click', function() {{
                        // Determine if this is a subcategory filter or investment type filter
                        const subcategory = this.getAttribute('data-subcategory');
                        const investmentType = this.getAttribute('data-investment-type');
                        
                        if (subcategory) {{
                            // Handle subcategory filtering
                            const isAll = subcategory === 'all';
                            
                            // Update active state for subcategory filters
                            segment.querySelectorAll('[data-subcategory]').forEach(btn => {{
                                btn.classList.remove('active');
                            }});
                            this.classList.add('active');
                            
                            // Filter companies by subcategory
                            segment.querySelectorAll('.company-card').forEach(card => {{
                                const cardSubcategory = card.querySelector('.subcategory-tag').textContent;
                                if (isAll || cardSubcategory === subcategory) {{
                                    card.style.display = '';
                                }} else {{
                                    card.style.display = 'none';
                                }}
                            }});
                        }} else if (investmentType) {{
                            // Handle investment type filtering
                            const isAll = investmentType === 'all';
                            
                            // Update active state for investment type filters
                            segment.querySelectorAll('[data-investment-type]').forEach(btn => {{
                                btn.classList.remove('active');
                            }});
                            this.classList.add('active');
                            
                            // Filter companies by investment type
                            segment.querySelectorAll('.company-card').forEach(card => {{
                                const cardType = card.getAttribute('data-investment-type');
                                if (isAll || cardType === investmentType) {{
                                    card.style.display = '';
                                }} else {{
                                    card.style.display = 'none';
                                }}
                            }});
                        }}
                        
                        // Hide investor sections with no visible companies
                        segment.querySelectorAll('.investor-section').forEach(investor => {{
                            const visibleCards = Array.from(investor.querySelectorAll('.company-card')).filter(card => 
                                card.style.display !== 'none'
                            );
                            
                            if (visibleCards.length === 0) {{
                                investor.style.display = 'none';
                            }} else {{
                                investor.style.display = '';
                            }}
                        }});
                    }});
                }});
            }}
        }});
    </script>
</body>
</html>
'''

# Save the HTML to a file
with open('bb_initiatives_overview.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

print("\nHTML report has been saved to 'bb_initiatives_overview.html'") 