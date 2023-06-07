import pandas as pd

arquivo_excel_5 = "INFORMATIVO 5° GRE.xlsx"
inf_5 = pd.read_excel(arquivo_excel_5, sheet_name="Sheet1")

for i in range(len(inf_5)):
    inf_5.loc[i, 'DIAS'] = inf_5['DIAS'][i][2:]
inf_5.to_excel('INFORMATIVO 5° GRE.xlsx', index=False)


arquivo_excel_18 = "INFORMATIVO 18° GRE.xlsx"
inf_18 = pd.read_excel(arquivo_excel_18, sheet_name="Sheet1")

for i in range(len(inf_18)):
    inf_18.loc[i, 'DIAS'] = inf_18['DIAS'][i][2:]
inf_18.to_excel('INFORMATIVO 18° GRE.xlsx', index=False)

print('\n\tPRONTINHO!\n')