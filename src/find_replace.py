
import uno
import string
import os, sys, csv, shutil

from faker import Factory
fake = Factory.create('pt_BR')


path = os.path.abspath('.')
path_input = path +"/input/"
path_output = path +"/output/"
path_subs = path +"/subs/"

subs = {}

output_txt_subs = []
output_txt_de_para = []

#carrega substituições
with open(path_subs+'input.csv') as csvfile:
    colunas = ['Arquivo', 'Nome', 'Sexo']
    reader = csv.DictReader(csvfile, fieldnames=colunas)
    for i, linha in enumerate(reader):
        if (linha['Nome'] != "") and (linha['Arquivo'] != "") and  (linha['Sexo'] != ""):
            subs[i] = {"Arquivo": linha['Arquivo'], "Nome": linha['Nome'], "Sexo": linha['Sexo']}


#Identifica o serviço do Libre Office
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
context = resolver.resolve("uno:socket,host=localhost,port=8100;urp;StarOffice.ComponentContext")
desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)

#carrega arquivos do diretório Input
dirs = os.listdir( path_input )

#copia todos os arquivos
for file in dirs:
    shutil.copy(path_input+file, path_output)

t = len(subs)

for i, (k,v) in enumerate(subs.items()):


    #v['Nome']
    #v['Arquivo']
    #v['Sexo']

    arquivo = v['Arquivo']+".doc"
    path_arquivo = "file://"+path_output+arquivo

    try:
        document = desktop.loadComponentFromURL( path_arquivo,"_blank", 0, ())
        print("(%d de %d) Arquivo %s encontrado " % (i,t,arquivo))

        search = document.createSearchDescriptor()
        cursor = document.Text.createTextCursor()

        search.SearchString = v['Nome']
        search.SearchCaseSensitive = True
        search.SearchWords = True

        found = document.findFirst( search )

        if found:
            if v['Sexo'] == "Masculino":
                 novo_nome = fake.name_male()
            else:
                 novo_nome = fake.name_female()

            output_txt_subs.append(arquivo+ " - " + v['Nome'])
            output_txt_de_para.append("Entrevista " + arquivo + " - DE: " + v['Nome'] + " PARA: " +  novo_nome)

        else:
           output_txt_subs.append("Entrevista " + arquivo + " - Ñão possui o nome " + v['Nome'])

        while found:
            found.String =  found.String.replace( v['Nome'],novo_nome)
            found = document.findNext( found.End, search)
            output_txt_subs.append(arquivo+ " - " + v['Nome'])

        path_arquivo_save = "file://"+path_output+arquivo
        document.storeAsURL(path_arquivo_save,())
        document.dispose()

    except:
        print("(%d de %d) Arquivo %s NAO encontrado " % (i,t,arquivo))

        output_txt_subs.append("Entrevista " + arquivo + " não encontrada.")


outFile = open(path_output+'resultado.txt', 'a')
outFile.write("Substituições Realizadas\n")
outFile.write("\n".join(output_txt_subs))
outFile.close()
outFile = open(path_output+'resultado_de_para.txt', 'a')
outFile.write("De-Para realizado\n")
outFile.write("\n".join(output_txt_de_para))
outFile.close()
