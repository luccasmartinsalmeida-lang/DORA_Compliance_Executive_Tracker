import pandas as pd

# Seed dataset simulating DORA Framework Requirements
data = [
    ["DORA-1.1", "ICT Risk Management", "ICT Risk Management Framework documented", "Fully Compliant", "Verified", "Low", "GRC & Security"],
    ["DORA-1.2", "ICT Risk Management", "Continuous asset mapping & identification", "Partially Compliant", "In Review", "Medium", "Infra & Sec"],
    ["DORA-1.3", "ICT Risk Management", "Annual review of ICT risk policies", "Fully Compliant", "Verified", "Low", "GRC & Security"],
    ["DORA-1.4", "ICT Risk Management", "Legacy system risk assessment", "Non-Compliant", "Missing", "Critical", "Enterprise Arch"],
    ["DORA-1.5", "ICT Risk Management", "Cryptographic controls & key management", "Partially Compliant", "In Review", "High", "Infra & Sec"],
    ["DORA-1.6", "ICT Risk Management", "Access control & identity management policy", "Fully Compliant", "Verified", "Low", "IAM Team"],
    ["DORA-1.7", "ICT Risk Management", "Multi-factor authentication (MFA) enforcement", "Fully Compliant", "Verified", "Low", "IAM Team"],
    ["DORA-1.8", "ICT Risk Management", "Network segmentation between prod/non-prod", "Partially Compliant", "In Review", "Medium", "Network Sec"],
    ["DORA-1.9", "ICT Risk Management", "Patch management SLA enforcement", "Non-Compliant", "Missing", "High", "Infra & Sec"],
    ["DORA-1.10", "ICT Risk Management", "Data loss prevention (DLP) monitoring", "Non-Compliant", "Missing", "High", "AppSec"],
    ["DORA-2.1", "ICT Incident Management", "ICT Incident classification methodology", "Fully Compliant", "Verified", "Low", "SOC Team"],
    ["DORA-2.2", "ICT Incident Management", "Major incident notification to regulator (<2h)", "Non-Compliant", "Missing", "Critical", "Regulatory Compliance"],
    ["DORA-2.3", "ICT Incident Management", "Centralized incident logging & SIEM integration", "Fully Compliant", "Verified", "Low", "SOC Team"],
    ["DORA-2.4", "ICT Incident Management", "Post-incident root cause analysis process", "Partially Compliant", "In Review", "Medium", "SOC Team"],
    ["DORA-2.5", "ICT Incident Management", "Crisis communication plan with stakeholders", "Fully Compliant", "Verified", "Low", "Corp Comms"],
    ["DORA-2.6", "ICT Incident Management", "Customer notification process for major incidents", "Partially Compliant", "In Review", "High", "Customer Operations"],
    ["DORA-2.7", "ICT Incident Management", "Incident response playbooks for ransomware", "Non-Compliant", "Missing", "Critical", "SOC Team"],
    ["DORA-2.8", "ICT Incident Management", "Anomalous behavior detection mechanisms", "Fully Compliant", "Verified", "Low", "SOC Team"],
    ["DORA-3.1", "Digital Operational Resilience Testing", "Annual vulnerability scanning program", "Fully Compliant", "Verified", "Low", "AppSec"],
    ["DORA-3.2", "Digital Operational Resilience Testing", "Threat-Led Penetration Testing (TLPT)", "Non-Compliant", "Missing", "Critical", "Red Team"],
    ["DORA-3.3", "Digital Operational Resilience Testing", "Independent testing of critical ICT tools", "Partially Compliant", "In Review", "High", "Audit Team"],
    ["DORA-3.4", "Digital Operational Resilience Testing", "Source code security review integration", "Fully Compliant", "Verified", "Low", "DevSecOps"],
    ["DORA-3.5", "Digital Operational Resilience Testing", "Disaster recovery plan dry-run testing", "Partially Compliant", "In Review", "Medium", "BCP Team"],
    ["DORA-3.6", "Digital Operational Resilience Testing", "Backup restoration verification testing", "Non-Compliant", "Missing", "Critical", "Infra & Sec"],
    ["DORA-3.7", "Digital Operational Resilience Testing", "Third-party penetration testing validation", "Partially Compliant", "In Review", "Medium", "TPRM Team"],
    ["DORA-3.8", "Digital Operational Resilience Testing", "Red team scenario exercises", "Non-Compliant", "Missing", "High", "Red Team"],
    ["DORA-4.1", "Third-Party Risk Management", "Register of Information (RoI) for ICT vendors", "Partially Compliant", "In Review", "High", "TPRM Team"],
    ["DORA-4.2", "Third-Party Risk Management", "Critical vendor concentration risk assessment", "Non-Compliant", "Missing", "Critical", "TPRM Team"],
    ["DORA-4.3", "Third-Party Risk Management", "Subcontracting chain risk monitoring", "Non-Compliant", "Missing", "High", "TPRM Team"],
    ["DORA-4.4", "Third-Party Risk Management", "Mandatory DORA clauses in vendor contracts", "Partially Compliant", "In Review", "High", "Legal & Procurement"],
    ["DORA-4.5", "Third-Party Risk Management", "Vendor exit strategy & transition plans", "Non-Compliant", "Missing", "Critical", "TPRM Team"],
    ["DORA-4.6", "Third-Party Risk Management", "Annual performance audit of critical vendors", "Fully Compliant", "Verified", "Low", "TPRM Team"],
    ["DORA-4.7", "Third-Party Risk Management", "SLA monitoring for critical ICT services", "Fully Compliant", "Verified", "Low", "Vendor Management"],
    ["DORA-4.8", "Third-Party Risk Management", "Vendor security questionnaire automation", "Partially Compliant", "In Review", "Medium", "TPRM Team"],
    ["DORA-5.1", "Information Sharing", "Threat intelligence sharing participation", "Fully Compliant", "Verified", "Low", "Threat Intel"],
    ["DORA-5.2", "Information Sharing", "Internal threat intelligence distribution", "Fully Compliant", "Verified", "Low", "Threat Intel"],
    ["DORA-5.3", "Information Sharing", "Peer-to-peer IoC exchange mechanism", "Partially Compliant", "In Review", "Medium", "Threat Intel"],
    ["DORA-5.4", "Information Sharing", "Regulatory cyber threat reporting", "Fully Compliant", "Verified", "Low", "Regulatory Compliance"]
]

cols = ["Requirement_ID", "Domain", "Requirement_Title", "Status", "Evidence_Status", "Risk_Level", "Owner_Team"]
df = pd.DataFrame(data, columns=cols)
df.to_csv('dora_compliance_master.csv', index=False)

output_file = "DORA_Compliance_Executive_Tracker.xlsx"

writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
workbook = writer.book

# Styling and Font Configurations
font_family = 'Calibri'

header_format = workbook.add_format({
    'bold': True, 'font_name': font_family, 'font_size': 11,
    'font_color': '#FFFFFF', 'bg_color': '#1E3A8A',
    'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D1D5DB'
})

header_red_format = workbook.add_format({
    'bold': True, 'font_name': font_family, 'font_size': 11,
    'font_color': '#FFFFFF', 'bg_color': '#991B1B',
    'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#D1D5DB'
})

border_format = workbook.add_format({'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})
border_center = workbook.add_format({'font_name': font_family, 'font_size': 10, 'align': 'center', 'border': 1, 'border_color': '#D1D5DB'})
percent_format = workbook.add_format({'font_name': font_family, 'font_size': 10, 'align': 'center', 'num_format': '0.0%', 'border': 1, 'border_color': '#D1D5DB'})

# Badge Formats for Compliance Status
fmt_fc = workbook.add_format({'bg_color': '#DCFCE7', 'font_color': '#15803D', 'bold': True, 'align': 'center', 'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})
fmt_pc = workbook.add_format({'bg_color': '#FEF3C7', 'font_color': '#B45309', 'bold': True, 'align': 'center', 'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})
fmt_nc = workbook.add_format({'bg_color': '#FEE2E2', 'font_color': '#B91C1C', 'bold': True, 'align': 'center', 'font_name': font_family, 'font_size': 10, 'border': 1, 'border_color': '#D1D5DB'})

# --- SHEET 1: EXECUTIVE SUMMARY ---
ws_dash = workbook.add_worksheet('Executive Summary')
ws_dash.hide_gridlines(2)

ws_dash.merge_range('A1:E1', "DORA COMPLIANCE & GAP ANALYSIS DASHBOARD", workbook.add_format({'bold': True, 'font_name': font_family, 'font_size': 14, 'font_color': '#1E3A8A'}))
ws_dash.merge_range('A2:E2', "EU Regulation (EU) 2022/2554 • Illustrative Benchmarking Model", workbook.add_format({'italic': True, 'font_name': font_family, 'font_size': 9, 'font_color': '#6B7280'}))

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

curr_r = 12
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

# Embedded Charts
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
    'name':       'Compliant',
    'categories': ['Executive Summary', 12, 0, curr_r - 1, 0],
    'values':     ['Executive Summary', 12, 1, curr_r - 1, 1],
})
bar.add_series({
    'name':       'Gaps',
    'categories': ['Executive Summary', 12, 0, curr_r - 1, 0],
    'values':     ['Executive Summary', 12, 2, curr_r - 1, 2],
})
bar.set_title({'name': 'Compliance Level by DORA Domain', 'name_font': {'name': font_family, 'size': 11, 'bold': True}})
bar.set_x_axis({'name': 'Domain', 'label_rotation': -20})
bar.set_y_axis({'name': 'Requirements'})
bar.set_legend({'position': 'right'})
bar.set_size({'width': 480, 'height': 280})
ws_dash.insert_chart('G18', bar)

ws_dash.set_column('A:A', 38)
ws_dash.set_column('B:E', 15)

# --- SHEET 2: ALL REQUIREMENTS ---
ws_all = workbook.add_worksheet('All Requirements')
ws_all.hide_gridlines(2)

for c_idx, h in enumerate(cols):
    ws_all.write(0, c_idx, h, header_format)

for r_idx, row in df.iterrows():
    for c_idx, val in enumerate(row):
        col_name = cols[c_idx]
        if col_name == 'Status':
            if val == 'Fully Compliant':
                ws_all.write(r_idx + 1, c_idx, val, fmt_fc)
            elif val == 'Partially Compliant':
                ws_all.write(r_idx + 1, c_idx, val, fmt_pc)
            else:
                ws_all.write(r_idx + 1, c_idx, val, fmt_nc)
        elif col_name in ['Requirement_ID', 'Evidence_Status', 'Risk_Level']:
            ws_all.write(r_idx + 1, c_idx, val, border_center)
        else:
            ws_all.write(r_idx + 1, c_idx, val, border_format)

ws_all.set_column('A:A', 18)
ws_all.set_column('B:B', 38)
ws_all.set_column('C:C', 50)
ws_all.set_column('D:G', 20)

# --- SHEET 3: URGENT REMEDIATION GAPS ---
ws_gaps = workbook.add_worksheet('Urgent Remediation Gaps')
ws_gaps.hide_gridlines(2)

urgent_df = df[(df['Status'] == 'Non-Compliant') & (df['Risk_Level'].isin(['Critical', 'High']))].reset_index(drop=True)

for c_idx, h in enumerate(cols):
    ws_gaps.write(0, c_idx, h, header_red_format)

for r_idx, row in urgent_df.iterrows():
    for c_idx, val in enumerate(row):
        col_name = cols[c_idx]
        if col_name == 'Status':
            ws_gaps.write(r_idx + 1, c_idx, val, fmt_nc)
        elif col_name in ['Requirement_ID', 'Evidence_Status', 'Risk_Level']:
            ws_gaps.write(r_idx + 1, c_idx, val, border_center)
        else:
            ws_gaps.write(r_idx + 1, c_idx, val, border_format)

ws_gaps.set_column('A:A', 18)
ws_gaps.set_column('B:B', 38)
ws_gaps.set_column('C:C', 50)
ws_gaps.set_column('D:G', 20)

# Close and export Excel File
writer.close()
print("✅ DORA Executive Tracker successfully exported to Excel with charts!")