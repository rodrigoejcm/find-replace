
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
    colunas = ['Arquivo', 'Nome', 'Primeiro_Nome', 'Sexo']
    reader = csv.DictReader(csvfile, fieldnames=colunas)
    for i, linha in enumerate(reader):
        if (linha['Nome'] != "") and (linha['Arquivo'] != "") and  (linha['Sexo'] != ""):
            subs[i] = {"Arquivo": linha['Arquivo'], "Nome": linha['Nome'], "Primeiro_Nome": linha['Primeiro_Nome'], "Sexo": linha['Sexo']}


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
    #v['Primeiro_Nome']
    #v['Arquivo']
    #v['Sexo']

    arquivo = v['Arquivo']+".doc"
    path_arquivo = "file://"+path_output+arquivo

    try:
        document = desktop.loadComponentFromURL( path_arquivo,"_blank", 0, ())
        print("(%d de %d) Arquivo %s encontrado " % (i,t,arquivo))

        search = document.createSearchDescriptor()
        cursor = document.Text.createTextCursor()


        #PRIMEIRO BUSCA O NOME COMPLETO
        search.SearchString = v['Nome']
        search.SearchCaseSensitive = True
        search.SearchWords = True

        found = document.findFirst( search )

        if found:
            if v['Sexo'] == "Masculino":
                 novo_nome = fake.first_name_male()
                 novo_nome_sobrenome = novo_nome + " " + fake.last_name()
            else:
                 novo_nome = fake.first_name_female()
                 novo_nome_sobrenome = novo_nome + " " + fake.last_name()

            output_txt_subs.append([ arquivo , "Susbstituição Nome Completo: " + v['Nome']])
            output_txt_de_para.append( [ arquivo,  v['Nome']  ,  novo_nome_sobrenome])

        else:
           output_txt_subs.append( [arquivo,  "Nome não encontrado: "  + v['Nome'] ] )

        while found:
            found.String =  found.String.replace( v['Nome'],novo_nome_sobrenome)
            found = document.findNext( found.End, search)
            output_txt_subs.append([ arquivo , "Susbstituição Nome Completo: " + v['Nome']])


        #DEPOIS BUSCA SOMENTO O PRIMEIRO NOME
        search.SearchString = v['Primeiro_Nome']
        search.SearchCaseSensitive = True
        search.SearchWords = True

        found = document.findFirst( search )

        if found:

            output_txt_subs.append([ arquivo , "Susbstituição Primeiro Nome: " + v['Primeiro_Nome']])
            output_txt_de_para.append( [ arquivo,  v['Primeiro_Nome']  ,  novo_nome])


        while found:
            found.String =  found.String.replace( v['Primeiro_Nome'] ,novo_nome)
            found = document.findNext( found.End, search)
            output_txt_subs.append([ arquivo , "Susbstituição Primeiro Nome: " + v['Primeiro_Nome'] ])

        path_arquivo_save = "file://"+path_output+arquivo
        document.storeAsURL(path_arquivo_save,())
        document.dispose()

    except:
        print("(%d de %d) Arquivo %s NAO encontrado " % (i,t,arquivo))
        output_txt_subs.append([ arquivo , "Arquivo não encontrado" ])


with open(path_output+'resultado_de_para.csv', 'w') as csvfile:
    campos = ['Arquivo', 'De _Nome', 'Para_Nome']
    writer = csv.DictWriter(csvfile, fieldnames=campos)
    writer.writeheader()

    for de_para in output_txt_de_para:
        writer.writerow({'Arquivo': de_para[0], 'De _Nome': de_para[1], 'Para_Nome' : de_para[2]})


with open(path_output+'resultado_substituicoes.csv', 'w') as csvfile:
    campos_2 = ['Arquivo', 'Resultado']
    writer = csv.DictWriter(csvfile, fieldnames=campos_2)
    writer.writeheader()

    for result in output_txt_subs:
        writer.writerow({ 'Arquivo': result[0], 'Resultado': result[1] })
