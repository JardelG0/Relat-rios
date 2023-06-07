import pandas as pd

# Rodando o informativo da 5° GRE
inf_5 = pd.read_excel("INFORMATIVO 5° GRE.xlsx", sheet_name="Sheet1")

try:
    for i in range(len(inf_5)):
        inf_5.loc[i, 'DIAS'] = inf_5['DIAS'][i][2:]
    inf_5.to_excel('INFORMATIVO 5° GRE.xlsx', index=False)
    print("\n\t ! INFORMATIVO 5° GRE ready !")
except:
    print('\n\t! INFORMATIVO 5° GRE already updated !\n')



# Rodando o informativo da 18° GRE
inf_18 = pd.read_excel("INFORMATIVO 18° GRE.xlsx", sheet_name="Sheet1")

try:
    for i in range(len(inf_18)):
        inf_18.loc[i, 'DIAS'] = inf_18['DIAS'][i][2:]
    inf_18.to_excel('INFORMATIVO 18° GRE.xlsx', index=False)
    print("\n\t ! INFORMATIVO 18° GRE ready !")
except:
    print('\n\t! INFORMATIVO 18° GRE already updated !\n')
