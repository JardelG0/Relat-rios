import pandas as pd

def main():
    perm = True
    while perm:
        alert = input("\n\t! ALERT !\n\nVocê tem certeza que deseja limpar os informativos? \nSim[S] Não[N]\n>_")
        if alert.upper() == "S":
            # Rodando o informativo da 5° GRE
            try:
                inf_5 = pd.read_excel("INFORMATIVO 5° GRE.xlsx", sheet_name="Sheet1")
                try:
                    cont_5 = len(inf_5)
                    for i in range(cont_5):
                        inf_5.loc[i, 'DIAS'] = inf_5['DIAS'][i][2:]

                    inf_5 = inf_5.T
                    for i in range(cont_5):
                        if inf_5[i][1] == "":
                            inf_5.pop(i)
                    inf_5 = inf_5.T

                    inf_5.to_excel('INFORMATIVO 5° GRE.xlsx', index=False)
                    print("\n\t ! INFORMATIVO 5° GRE ready !")
                except:
                    print('\n\t! INFORMATIVO 5° GRE already updated !\n')
            except:
                print('\n\tINFORMATIVO NOT FOUND!')


            # Rodando o informativo da 16° GRE
            try:
                inf_16 = pd.read_excel("INFORMATIVO 16° GRE.xlsx", sheet_name="Sheet1")
                try:
                    cont_16 = len(inf_16)
                    for i in range(cont_16):
                        inf_16.loc[i, 'DIAS'] = inf_16['DIAS'][i][2:]
                    
                    inf_16 = inf_16.T
                    for i in range(cont_16):
                        if inf_16[i][1] == "":
                            inf_16.pop(i)
                    inf_16 = inf_16.T

                    inf_16.to_excel('INFORMATIVO 16° GRE.xlsx', index=False)
                    print("\n\t ! INFORMATIVO 16° GRE ready !")
                except:
                    print('\n\t! INFORMATIVO 16° GRE already updated !\n')
            except:
                print('\n\tINFORMATIVO NOT FOUND!')


            # Rodando o informativo da 18° GRE
            try:
                inf_18 = pd.read_excel("INFORMATIVO 18° GRE.xlsx", sheet_name="Sheet1")
                try:
                    cont_18 = len(inf_18)
                    for i in range(cont_18):
                        inf_18.loc[i, 'DIAS'] = inf_18['DIAS'][i][2:]
                    
                    inf_18 = inf_18.T
                    for i in range(cont_18):
                        if inf_18[i][1] == "":
                            inf_18.pop(i)
                    inf_18 = inf_18.T

                    inf_18.to_excel('INFORMATIVO 18° GRE.xlsx', index=False)
                    print("\n\t ! INFORMATIVO 18° GRE ready !")
                    perm = False
                except:
                    print('\n\t! INFORMATIVO 18° GRE already updated !\n')
            except:
                print('\n\tINFORMATIVO NOT FOUND!')

            perm = False
        elif alert.upper() == "N": 
            print("\n\t! OPERAÇÃO ABORTADA !\n")
            perm = False
        else:
            print("Valor Inválido")

if __name__ == "__main__":
    main()