import os
import creds

diretorio = creds.user_path
listaPastas = []
listaArquivos = []

provisorio = os.chdir(diretorio)
listaPastas.append(os.listdir(provisorio))

for a in listaPastas[0]:
    provisorio2 = os.chdir(diretorio + a)

    for b in os.listdir(provisorio2):
        listaArquivos.append(b)

# print(listaPastas)
print(listaArquivos)