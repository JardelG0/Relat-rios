import os

def main():
    perm = True
    while perm:
        alert = input("\n\t! ALERT !\n\nVocê tem certeza que deseja limpar os diretórios? \nSim[S] Não[N]\n>_")

        if alert.upper() == "S":
            gres = ['\\5° GRE\\', '\\16° GRE\\', '\\18° GRE\\']
            dir_1 = os.getcwd()
            for i in range(len(gres)):
                os.chdir(dir_1 + gres[i])
                dir_2 = os.getcwd()
                for j in os.listdir():
                    if j[-5:] != ".docx":
                        os.chdir(os.getcwd() + "\\" + j)
                        for k in os.listdir():
                            os.remove(k)
                        os.chdir(dir_2)
            print("\n\n\t! Cleaned Directories ! ")
            perm = False
        elif alert.upper() == "N": 
            print("\n\t! OPERAÇÃO ABORTADA !\n")
            perm = False
        else:
            print("Valor Inválido")


if __name__ == "__main__":
    main()