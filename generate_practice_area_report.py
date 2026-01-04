import pandas as pd
import csv

# Read raw data
df = pd.read_csv('/home/user/Reports/Weekly Intake_Sales Tracker Fixed Formulas 2025.xlsx - December Data.csv')

# Clean revenue column
df['Case Revenue ($)'] = df['Case Revenue ($)'].replace(r'[\$,]', '', regex=True)
df['Case Revenue ($)'] = pd.to_numeric(df['Case Revenue ($)'], errors='coerce').fillna(0)

# Identify strategy sessions
df['Is_Strategy_Session'] = df['Additional Notes'].str.contains('Strategy Session', case=False, na=False)

# Define practice area groupings
def map_practice_area(type_of_law):
    if pd.isna(type_of_law):
        return None
    type_of_law = str(type_of_law).strip()
    if type_of_law == 'Family':
        return 'Family Law'
    elif type_of_law in ['Criminal', 'Domestic Violence']:
        return 'Criminal Law'
    elif type_of_law == 'Civil':
        return 'Civil Law'
    elif type_of_law == 'Real Estate':
        return 'Real Estate'
    elif type_of_law in ['Estate Planning', 'Probate']:
        return 'Estate Law'
    else:
        return None  # Exclude 'Other' and unmapped

df['Practice Area'] = df['Type of Law'].apply(map_practice_area)

# Filter to only mapped practice areas
df_filtered = df[df['Practice Area'].notna()].copy()

# Practice areas in order
practice_areas = ['Family Law', 'Criminal Law', 'Civil Law', 'Real Estate', 'Estate Law']

def fmt_currency(val):
    if val == 0:
        return "$0"
    return f"${val:,.0f}"

def get_metrics(pa_df):
    """Get metrics for a practice area"""
    leads = len(pa_df)
    mqls = len(pa_df[pa_df['Marketing Qualified Lead (MQL)?'] == 'Yes'])
    sqls = len(pa_df[pa_df['Sales Qualified Lead (SQL)?'] == 'Yes'])
    scheduled = len(pa_df[pa_df['Consult Scheduled?'] == 'Yes'])
    showups = len(pa_df[(pa_df['Consult Scheduled?'] == 'Yes') & (pa_df['No Show (Yes/No)'] == 'No')])
    hired_all = len(pa_df[pa_df['Sales Outcome'] == 'Hired'])
    strategy = len(pa_df[(pa_df['Sales Outcome'] == 'Hired') & (pa_df['Is_Strategy_Session'])])
    new_matters = hired_all - strategy  # Hired excluding strategy sessions
    mkt_contacts = len(pa_df[pa_df['Marketing Contacts'].notna() & (pa_df['Marketing Contacts'] != '')])
    revenue = pa_df['Case Revenue ($)'].sum()
    return {
        'Leads': leads,
        'Qualified Leads (MQLs)': mqls,
        'Scheduled': scheduled,
        'Showups': showups,
        'New Matters': new_matters,
        'Marketing Contacts': mkt_contacts,
        'SQLs': sqls,
        'Monthly Total Matter Value': revenue
    }

def get_source_breakdown(pa_df):
    """Get source breakdown for a practice area"""
    sources = []
    for source_name, group in pa_df.groupby('Lead Source'):
        leads = len(group)
        hired = len(group[group['Sales Outcome'] == 'Hired'])
        revenue = group['Case Revenue ($)'].sum()
        sources.append({
            'Source': source_name,
            'Total Leads': leads,
            'Total Hires': hired,
            'Total Matter Value': revenue
        })
    # Sort by Total Leads descending
    sources = sorted(sources, key=lambda x: x['Total Leads'], reverse=True)
    return sources

# Generate output
output_rows = []

for pa in practice_areas:
    pa_df = df_filtered[df_filtered['Practice Area'] == pa]

    if len(pa_df) == 0:
        continue

    # Practice Area Header
    output_rows.append([pa])
    output_rows.append([])  # Empty row

    # Metrics Table Header
    output_rows.append(['Metrics Summary'])
    metrics = get_metrics(pa_df)
    output_rows.append(['Metric', 'Value'])
    output_rows.append(['Leads', metrics['Leads']])
    output_rows.append(['Qualified Leads (MQLs)', metrics['Qualified Leads (MQLs)']])
    output_rows.append(['Scheduled', metrics['Scheduled']])
    output_rows.append(['Showups', metrics['Showups']])
    output_rows.append(['New Matters', metrics['New Matters']])
    output_rows.append(['Marketing Contacts', metrics['Marketing Contacts']])
    output_rows.append(['SQLs', metrics['SQLs']])
    output_rows.append(['Monthly Total Matter Value', fmt_currency(metrics['Monthly Total Matter Value'])])

    output_rows.append([])  # Empty row

    # Source Breakdown Table
    output_rows.append(['Source Breakdown'])
    output_rows.append(['Source', 'Total Leads', 'Total Hires', 'Total Matter Value'])

    source_data = get_source_breakdown(pa_df)
    total_leads = 0
    total_hires = 0
    total_value = 0

    for s in source_data:
        output_rows.append([s['Source'], s['Total Leads'], s['Total Hires'], fmt_currency(s['Total Matter Value'])])
        total_leads += s['Total Leads']
        total_hires += s['Total Hires']
        total_value += s['Total Matter Value']

    # Total row
    output_rows.append(['TOTAL', total_leads, total_hires, fmt_currency(total_value)])

    output_rows.append([])  # Empty row
    output_rows.append([])  # Extra spacing between practice areas
    output_rows.append([])

# Write to CSV
output_file = '/home/user/Reports/December_Practice_Area_Breakdown.csv'
with open(output_file, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows(output_rows)

print(f"Report generated: {output_file}")
print(f"\nSummary by Practice Area:")
for pa in practice_areas:
    pa_df = df_filtered[df_filtered['Practice Area'] == pa]
    if len(pa_df) > 0:
        metrics = get_metrics(pa_df)
        print(f"\n{pa}:")
        print(f"  Leads: {metrics['Leads']}")
        print(f"  New Matters: {metrics['New Matters']}")
        print(f"  Revenue: {fmt_currency(metrics['Monthly Total Matter Value'])}")
