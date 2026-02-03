import pandas as pd
import csv

# Read raw data
df = pd.read_csv('/home/user/Reports/Master Monthly Marketing Tracker 2025 (Charles Luna Sheet).xlsx - Weekly (1_22-1_27).csv')

# Clean revenue column
df['Case Revenue ($)'] = df['Case Revenue ($)'].replace(r'[\$,]', '', regex=True)
df['Case Revenue ($)'] = pd.to_numeric(df['Case Revenue ($)'], errors='coerce').fillna(0)

# Identify strategy sessions (case-insensitive search for "Strategy Session")
df['Is_Strategy_Session'] = df['Additional Notes'].str.contains('Strategy Session', case=False, na=False)

# === SUMMARY METRICS ===
total_leads = len(df)
total_hired = len(df[df['Sales Outcome'] == 'Hired'])
strategy_sessions_count = len(df[(df['Sales Outcome'] == 'Hired') & (df['Is_Strategy_Session'])])
total_hired_excl_strategy = total_hired - strategy_sessions_count
total_sqls = len(df[df['Sales Qualified Lead (SQL)?'] == 'Yes'])
total_mqls = len(df[df['Marketing Qualified Lead (MQL)?'] == 'Yes'])
total_scheduled = len(df[df['Consult Scheduled?'] == 'Yes'])
total_showups = len(df[(df['Consult Scheduled?'] == 'Yes') & (df['No Show (Yes/No)'] == 'No')])
total_no_shows = len(df[(df['Consult Scheduled?'] == 'Yes') & (df['No Show (Yes/No)'] == 'Yes')])
total_marketing_contacts = len(df[df['Marketing Contacts'].notna() & (df['Marketing Contacts'] != '')])
avg_daily_leads = round(total_leads / 7, 1)
consultation_rate = round((total_scheduled / total_leads) * 100, 0) if total_leads > 0 else 0
conversion_rate = round((total_hired / total_leads) * 100, 2) if total_leads > 0 else 0
mql_rate = round((total_mqls / total_leads) * 100, 2) if total_leads > 0 else 0
sql_rate = round((total_sqls / total_leads) * 100, 2) if total_leads > 0 else 0
total_revenue = df['Case Revenue ($)'].sum()

# === INTAKE PERFORMANCE ===
def get_agent_metrics(group):
    leads = len(group)
    mqls = len(group[group['Marketing Qualified Lead (MQL)?'] == 'Yes'])
    sqls = len(group[group['Sales Qualified Lead (SQL)?'] == 'Yes'])
    scheduled = len(group[group['Consult Scheduled?'] == 'Yes'])
    showups = len(group[(group['Consult Scheduled?'] == 'Yes') & (group['No Show (Yes/No)'] == 'No')])
    no_shows = len(group[(group['Consult Scheduled?'] == 'Yes') & (group['No Show (Yes/No)'] == 'Yes')])
    hired_all = len(group[group['Sales Outcome'] == 'Hired'])
    strategy = len(group[(group['Sales Outcome'] == 'Hired') & (group['Is_Strategy_Session'])])
    hired_excl = hired_all - strategy
    revenue = group['Case Revenue ($)'].sum()
    return {
        'Total Leads': leads,
        'MQLs': mqls,
        'SQLs': sqls,
        'Scheduled': scheduled,
        'Showups': showups,
        'No Shows': no_shows,
        'Hired (All)': hired_all,
        'Hired (Excl. Strategy)': hired_excl,
        'Strategy Sessions': strategy,
        'Converted Sessions': strategy,
        'Unconverted Sessions': 0,
        'Total Value': revenue
    }

agent_data = []
for agent_name, group in df.groupby('Intake Agent Name'):
    metrics = get_agent_metrics(group)
    metrics['Agent'] = agent_name
    agent_data.append(metrics)

# Sort by Total Leads descending
agent_data = sorted(agent_data, key=lambda x: x['Total Leads'], reverse=True)

# Calculate totals for agents
agent_totals = {
    'Agent': 'TOTAL',
    'Total Leads': total_leads,
    'MQLs': total_mqls,
    'SQLs': total_sqls,
    'Scheduled': total_scheduled,
    'Showups': total_showups,
    'No Shows': total_no_shows,
    'Hired (All)': total_hired,
    'Hired (Excl. Strategy)': total_hired_excl_strategy,
    'Strategy Sessions': strategy_sessions_count,
    'Converted Sessions': strategy_sessions_count,
    'Unconverted Sessions': 0,
    'Total Value': total_revenue
}

# === REASON ANALYSIS ===
reason_counts = df['Reason Not Scheduled'].value_counts()
reason_data = []
for reason, count in reason_counts.items():
    pct = round((count / total_leads) * 100, 1)
    reason_data.append({'Reason': reason, 'Count': count, 'Percentage': pct})

# === PRACTICE AREA BREAKDOWN ===
practice_areas = ['Family', 'Civil', 'Criminal', 'Real Estate', 'Estate Planning', 'Probate']

def get_practice_metrics(pa_df):
    leads = len(pa_df)
    mqls = len(pa_df[pa_df['Marketing Qualified Lead (MQL)?'] == 'Yes'])
    sqls = len(pa_df[pa_df['Sales Qualified Lead (SQL)?'] == 'Yes'])
    scheduled = len(pa_df[pa_df['Consult Scheduled?'] == 'Yes'])
    showups = len(pa_df[(pa_df['Consult Scheduled?'] == 'Yes') & (pa_df['No Show (Yes/No)'] == 'No')])
    hired_all = len(pa_df[pa_df['Sales Outcome'] == 'Hired'])
    strategy = len(pa_df[(pa_df['Sales Outcome'] == 'Hired') & (pa_df['Is_Strategy_Session'])])
    hired_excl = hired_all - strategy
    mkt_contacts = len(pa_df[pa_df['Marketing Contacts'].notna() & (pa_df['Marketing Contacts'] != '')])
    revenue = pa_df['Case Revenue ($)'].sum()
    return {
        'Leads': leads,
        'MQLs': mqls,
        'SQLs': sqls,
        'Scheduled': scheduled,
        'Showups': showups,
        'New Matters': hired_excl,
        'Marketing Contacts': mkt_contacts,
        'Total Value': revenue,
        'Strategy Sessions': strategy,
        'Converted Sessions': strategy
    }

practice_data = {}
for pa in practice_areas:
    pa_df = df[df['Type of Law'] == pa]
    practice_data[pa] = get_practice_metrics(pa_df)

# Other category
other_df = df[~df['Type of Law'].isin(practice_areas)]
practice_data['Other'] = get_practice_metrics(other_df)

# === SOURCE ANALYSIS ===
def get_source_metrics(group):
    leads = len(group)
    mqls = len(group[group['Marketing Qualified Lead (MQL)?'] == 'Yes'])
    sqls = len(group[group['Sales Qualified Lead (SQL)?'] == 'Yes'])
    scheduled = len(group[group['Consult Scheduled?'] == 'Yes'])
    hired = len(group[group['Sales Outcome'] == 'Hired'])
    revenue = group['Case Revenue ($)'].sum()
    conv_rate = round((hired / leads) * 100, 1) if leads > 0 else 0
    sched_rate = round((scheduled / leads) * 100, 1) if leads > 0 else 0
    mql_rate_src = round((mqls / leads) * 100, 1) if leads > 0 else 0
    sql_rate_src = round((sqls / leads) * 100, 1) if leads > 0 else 0
    return {
        'Total Leads': leads,
        'MQLs': mqls,
        'SQLs': sqls,
        'Scheduled': scheduled,
        'Hired': hired,
        'Total Value': revenue,
        'Conversion Rate (%)': conv_rate,
        'Schedule Rate (%)': sched_rate,
        'MQL Rate (%)': mql_rate_src,
        'SQL Rate (%)': sql_rate_src
    }

source_data = []
for source_name, group in df.groupby('Lead Source'):
    metrics = get_source_metrics(group)
    metrics['Channel'] = source_name
    source_data.append(metrics)

# Sort by Total Leads descending
source_data = sorted(source_data, key=lambda x: x['Total Leads'], reverse=True)

# Source totals
source_totals = {
    'Channel': 'TOTAL',
    'Total Leads': total_leads,
    'MQLs': total_mqls,
    'SQLs': total_sqls,
    'Scheduled': total_scheduled,
    'Hired': total_hired,
    'Total Value': total_revenue,
    'Conversion Rate (%)': conversion_rate,
    'Schedule Rate (%)': round((total_scheduled / total_leads) * 100, 1) if total_leads > 0 else 0,
    'MQL Rate (%)': mql_rate,
    'SQL Rate (%)': sql_rate
}

# === FORMAT CURRENCY ===
def fmt_currency(val):
    if val == 0:
        return "$0"
    return f"${val:,.0f}"

# === GENERATE OUTPUT CSV ===
output_rows = []

# Row 1: Headers
row1 = ['Summary Table', '', '', '', 'Intake Performance'] + [''] * 12
output_rows.append(row1)

# Row 2: Metric/Value headers + Agent headers
row2 = ['Metric', 'Value', '', '', 'Agent', 'Total Leads', 'MQLs', 'SQLs', 'Scheduled', 'Showups',
        'No Shows', 'Hired (All)', 'Hired (Excl. Strategy)', 'Strategy Sessions', 'Converted Sessions',
        'Unconverted Sessions', 'Total Value']
output_rows.append(row2)

# Row 3-9: Summary metrics + Agent data
summary_metrics = [
    ('Total Leads', total_leads),
    ('Total Hired (All)', total_hired),
    ('Total Hired (Excluding Strategy Sessions)', total_hired_excl_strategy),
    ('Total SQLs', total_sqls),
    ('Total MQLs', total_mqls),
    ('Total Scheduled', total_scheduled),
    ('Total Showups', total_showups),
]

for i, (metric, value) in enumerate(summary_metrics):
    row = [metric, value, '', '']
    if i < len(agent_data):
        a = agent_data[i]
        row += [a['Agent'], a['Total Leads'], a['MQLs'], a['SQLs'], a['Scheduled'], a['Showups'],
                a['No Shows'], a['Hired (All)'], a['Hired (Excl. Strategy)'], a['Strategy Sessions'],
                a['Converted Sessions'], a['Unconverted Sessions'], fmt_currency(a['Total Value'])]
    else:
        row += [''] * 13
    output_rows.append(row)

# Row 10: Total No Shows + TOTAL row for agents
row10 = ['Total No Shows', total_no_shows, '', '',
         agent_totals['Agent'], agent_totals['Total Leads'], agent_totals['MQLs'], agent_totals['SQLs'],
         agent_totals['Scheduled'], agent_totals['Showups'], agent_totals['No Shows'],
         agent_totals['Hired (All)'], agent_totals['Hired (Excl. Strategy)'],
         agent_totals['Strategy Sessions'], agent_totals['Converted Sessions'],
         agent_totals['Unconverted Sessions'], fmt_currency(agent_totals['Total Value'])]
output_rows.append(row10)

# Row 11-17: More summary metrics
more_metrics = [
    ('Total Marketing Contacts', total_marketing_contacts),
    ('Average Daily Leads', avg_daily_leads),
    ('Consultation Rate (%)', f"{int(consultation_rate)}%"),
    ('Conversion Rate (%)', f"{conversion_rate}%"),
    ('MQL Rate (%)', f"{mql_rate}%"),
    ('SQL Rate (%)', f"{sql_rate}%"),
    ('Total Revenue', fmt_currency(total_revenue)),
]

for metric, value in more_metrics:
    row = [metric, value, '', ''] + [''] * 13
    output_rows.append(row)

# Rows 18-20: Empty
for _ in range(3):
    output_rows.append([''] * 17)

# Row 21: Reason Analysis header
row21 = ['Reason Analysis'] + [''] * 16
output_rows.append(row21)

# Row 22: Reason headers
row22 = ['Reason', 'Count', 'Percentage', ''] + [''] * 13
output_rows.append(row22)

# Practice area order for side-by-side display
pa_order = ['Family', 'Civil', 'Criminal', 'Real Estate', 'Probate', 'Other']
pa_headers = ['Leads', 'MQLs', 'SQLs', 'Scheduled', 'Showups', 'New Matters', 'Marketing Contacts',
              'Total Value', 'Strategy Sessions', 'Converted Sessions']

# Row 23+: Reason data + Practice Area data
pa_idx = 0
pa_row_count = 0
current_pa = None

for i, reason_row in enumerate(reason_data):
    row = [reason_row['Reason'], reason_row['Count'], reason_row['Percentage'], '']

    # Add practice area data on same rows
    if pa_idx < len(pa_order):
        if pa_row_count == 0:
            # Practice area name row
            current_pa = pa_order[pa_idx]
            row += [current_pa] + [''] * 12
        elif pa_row_count == 1:
            # Headers row
            row += pa_headers + [''] * 3
        elif pa_row_count == 2:
            # Data row
            pa_metrics = practice_data[current_pa]
            row += [pa_metrics['Leads'], pa_metrics['MQLs'], pa_metrics['SQLs'], pa_metrics['Scheduled'],
                    pa_metrics['Showups'], pa_metrics['New Matters'], pa_metrics['Marketing Contacts'],
                    fmt_currency(pa_metrics['Total Value']), pa_metrics['Strategy Sessions'],
                    pa_metrics['Converted Sessions']] + [''] * 3
        else:
            row += [''] * 13

        pa_row_count += 1
        if pa_row_count >= 5:  # 5 rows per practice area (name, headers, data, 2 blank)
            pa_row_count = 0
            pa_idx += 1
    else:
        row += [''] * 13

    output_rows.append(row)

# Continue with remaining practice areas if reason list was shorter
while pa_idx < len(pa_order):
    row = ['', '', '', '']
    if pa_row_count == 0:
        current_pa = pa_order[pa_idx]
        row += [current_pa] + [''] * 12
    elif pa_row_count == 1:
        row += pa_headers + [''] * 3
    elif pa_row_count == 2:
        pa_metrics = practice_data[current_pa]
        row += [pa_metrics['Leads'], pa_metrics['MQLs'], pa_metrics['SQLs'], pa_metrics['Scheduled'],
                pa_metrics['Showups'], pa_metrics['New Matters'], pa_metrics['Marketing Contacts'],
                fmt_currency(pa_metrics['Total Value']), pa_metrics['Strategy Sessions'],
                pa_metrics['Converted Sessions']] + [''] * 3
    else:
        row += [''] * 13

    pa_row_count += 1
    if pa_row_count >= 5:
        pa_row_count = 0
        pa_idx += 1

    output_rows.append(row)

# Add some blank rows before source analysis
for _ in range(3):
    output_rows.append([''] * 17)

# Source Analysis header
output_rows.append(['', '', '', '', 'Source Analysis'] + [''] * 12)

# Source Analysis column headers
source_headers = ['', '', '', '', 'Channel', 'Total Leads', 'MQLs', 'SQLs', 'Scheduled', 'Hired',
                  'Total Value', 'Conversion Rate (%)', 'Schedule Rate (%)', 'MQL Rate (%)', 'SQL Rate (%)'] + [''] * 2
output_rows.append(source_headers)

# Source data rows
for s in source_data:
    row = ['', '', '', '', s['Channel'], s['Total Leads'], s['MQLs'], s['SQLs'], s['Scheduled'],
           s['Hired'], fmt_currency(s['Total Value']), s['Conversion Rate (%)'], s['Schedule Rate (%)'],
           s['MQL Rate (%)'], s['SQL Rate (%)']] + [''] * 2
    output_rows.append(row)

# Source totals row
row = ['', '', '', '', source_totals['Channel'], source_totals['Total Leads'], source_totals['MQLs'],
       source_totals['SQLs'], source_totals['Scheduled'], source_totals['Hired'],
       fmt_currency(source_totals['Total Value']), source_totals['Conversion Rate (%)'],
       source_totals['Schedule Rate (%)'], source_totals['MQL Rate (%)'], source_totals['SQL Rate (%)']] + [''] * 2
output_rows.append(row)

# Write to CSV
output_file = '/home/user/Reports/Weekly (01_22-01_27) - Organized.csv'
with open(output_file, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerows(output_rows)

print(f"Report generated: {output_file}")
print(f"\nSummary:")
print(f"  Total Leads: {total_leads}")
print(f"  Total Hired: {total_hired}")
print(f"  Total Revenue: {fmt_currency(total_revenue)}")
print(f"  Strategy Sessions: {strategy_sessions_count}")
