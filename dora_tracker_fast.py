import os
import sys
import pandas as pd

df = pd.read_csv('dora_compliance_master.csv')
df.columns = df.columns.str.strip()

summary = df.groupby('Domain')['Status'].value_counts().unstack(fill_value=0)
summary['Total'] = summary.sum(axis=1)

# Validação direta sem estragar o broadcast do Pandas se 'Fully Compliant' não existir
fc = summary['Fully Compliant'] if 'Fully Compliant' in summary.columns else 0
summary['Compliance_Rate'] = (fc / summary['Total']).round(4)

urgent = df[(df['Status'] == 'Non-Compliant') & (df['Risk_Level'].isin(['Critical', 'High']))]

out = 'dora_compliance_tracker.xlsx'
with pd.ExcelWriter(out) as writer:
    summary.to_excel(writer, sheet_name='Summary')
    df.to_excel(writer, sheet_name='All_Data', index=False)
    urgent.to_excel(writer, sheet_name='Urgent_Gaps', index=False)

if sys.platform == 'win32':
    os.startfile(out)
elif sys.platform == 'darwin':
    os.system(f'open "{out}"')