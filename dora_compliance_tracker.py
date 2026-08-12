import pandas as pd

# 1. Load the raw data directly from the existing Excel file
input_file = "DORA_Raw_Data_Dashboard.xlsx"
df = pd.read_excel(input_file, sheet_name='Raw Data')

# 2. Setup the new Executive Dashboard output file
output_file = "DORA_Executive_Dashboard_Final.xlsx"
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

font_family = 'Calibri'

header_format = workbook.add_format({
    'bold': True, 'font_name': font_family, 'font_size': 11,
    'font_color': '#FFFFFF', 'bg_color': '#1E3A8A',
    'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D1D5DB'
})

border_format = workbook.add_format({'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})
border_center = workbook.add_format({'font_name': font_family, 'font_size': 10, 'align': 'center', 'border': 1, 'border_color': '#D1D5DB'})
percent_format = workbook.add_format({'font_name': font_family, 'font_size': 10, 'align': 'center', 'num_format': '0.0%', 'border': 1, 'border_color': '#D1D5DB'})

fmt_fc = workbook.add_format({'bg_color': '#DCFCE7', 'font_color': '#15803D', 'bold': True, 'align': 'center', 'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})
fmt_pc = workbook.add_format({'bg_color': '#FEF3C7', 'font_color': '#B45309', 'bold': True, 'align': 'center', 'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})
fmt_nc = workbook.add_format({'bg_color': '#FEE2E2', 'font_color': '#B91C1C', 'bold': True, 'align': 'center', 'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})

cols = df.columns.tolist()

# Sheet 1: Raw Data
ws_raw = workbook.add_worksheet('Raw Data')
ws_raw.hide_gridlines(2)

for c_idx, h in enumerate(cols):
    ws_raw.write(0, c_idx, h, header_format)

for r_idx, row in df.iterrows():
    for c_idx, val in enumerate(row):
        col_name = cols[c_idx]
        if col_name == 'Status':
            if val == 'Fully Compliant':
                ws_raw.write(r_idx + 1, c_idx, val, fmt_fc)
            elif val == 'Partially Compliant':
                ws_raw.write(r_idx + 1, c_idx, val, fmt_pc)
            else:
                ws_raw.write(r_idx + 1, c_idx, val, fmt_nc)
        elif col_name in ['Requirement_ID', 'Evidence_Status', 'Risk_Level']:
            ws_raw.write(r_idx + 1, c_idx, val, border_center)
        else:
            ws_raw.write(r_idx + 1, c_idx, val, border_format)

ws_raw.set_column('A:A', 18)
ws_raw.set_column('B:B', 38)
ws_raw.set_column('C:C', 50)
ws_raw.set_column('D:G', 20)

# Sheet 2: Executive Summary Dashboard
ws_dash = workbook.add_worksheet('Executive Summary')
ws_dash.hide_gridlines(2)

ws_dash.merge_range('A1:E1', "DORA COMPLIANCE - EXECUTIVE DASHBOARD", workbook.add_format({'bold': True, 'font_name': font_family, 'font_size': 14, 'font_color': '#1E3A8A'}))
ws_dash.merge_range('A2:E2', "EU Regulation (EU) 2022/2554 • Dynamic Data Integration", workbook.add_format({'italic': True, 'font_name': font_family, 'font_size': 9, 'font_color': '#6B7280'}))

total_reqs = len(df)
fully_comp = len(df[df['Status'] == 'Fully Compliant'])
part_comp = len(df[df['Status'] == 'Partially Compliant'])
non_comp = len(df[df['Status'] == 'Non-Compliant'])
compliance_rate = round((fully_comp / total_reqs), 3)

ws_dash.write('A4', 'Metric / Indicator', header_format)
ws_dash.write('B4', 'Count', header_format)

metrics = [("Fully Compliant", fully_comp), ("Partially Compliant", part_comp), ("Non-Compliant", non_comp), ("Total Evaluated", total_reqs)]
for idx, (label, val) in enumerate(metrics, start=4):
    ws_dash.write(idx, 0, label, border_format)
    ws_dash.write(idx, 1, val, border_center)

ws_dash.write('A9', 'Overall Compliance Rate', workbook.add_format({'bold': True, 'font_name': font_family, 'border': 1, 'border_color': '#D1D5DB'}))
ws_dash.write('B9', compliance_rate, percent_format)

ws_dash.write('A12', 'Domain Summary', header_format)
ws_dash.write('B12', 'Compliant', header_format)
ws_dash.write('C12', 'Gaps', header_format)
ws_dash.write('D12', 'Total', header_format)
ws_dash.write('E12', '% Compliance', header_format)

curr_r = 13
for dom in df['Domain'].unique():
    dom_sub = df[df['Domain'] == dom]
    d_tot = len(dom_sub)
    d_comp = len(dom_sub[dom_sub['Status'] == 'Fully Compliant'])
    d_gaps = len(dom_sub[dom_sub['Status'] != 'Fully Compliant'])
    d_rate = d_comp / d_tot
    
    ws_dash.write(curr_r, 0, dom, border_format)
    ws_dash.write(curr_r, 1, d_comp, border_center)
    ws_dash.write(curr_r, 2, d_gaps, border_center)
    ws_dash.write(curr_r, 3, d_tot, border_center)
    ws_dash.write(curr_r, 4, d_rate, percent_format)
    curr_r += 1

pie = workbook.add_chart({'type': 'pie'})
pie.add_series({
    'name': 'Status Distribution',
    'categories': ['Executive Summary', 4, 0, 6, 0],
    'values':     ['Executive Summary', 4, 1, 6, 1],
})
pie.set_title({'name': 'Status Distribution', 'name_font': {'name': font_family, 'size': 11, 'bold': True}})
pie.set_size({'width': 350, 'height': 240})
ws_dash.insert_chart('G4', pie)

bar = workbook.add_chart({'type': 'column'})
bar.add_series({
    'name':        'Compliant',
    'categories': ['Executive Summary', 13, 0, curr_r - 1, 0],
    'values':     ['Executive Summary', 13, 1, curr_r - 1, 1],
})
bar.add_series({
    'name':        'Gaps',
    'categories': ['Executive Summary', 13, 0, curr_r - 1, 0],
    'values':     ['Executive Summary', 13, 2, curr_r - 1, 2],
})
bar.set_title({'name': 'Compliance Level by Domain', 'name_font': {'name': font_family, 'size': 11, 'bold': True}})
bar.set_x_axis({'name': 'Domain', 'label_rotation': -20})
bar.set_y_axis({'name': 'Requirements'})
bar.set_legend({'position': 'right'})
bar.set_size({'width': 480, 'height': 280})
ws_dash.insert_chart('G18', bar)

ws_dash.set_column('A:A', 38)
ws_dash.set_column('B:E', 15)

writer.close()
print("Dashboard successfully generated from external file!")python dora_compliance_tracker.py