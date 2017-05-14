# Find-Replace
Scritp utilizado para desidentificação de múltiplos arquivos texto

## Dependências: 
  - Python 3
  - LibreOffice
  - UNO Tools https://pypi.python.org/pypi/unotools 0.3.3
  - Faker https://pypi.python.org/pypi/Faker  0.7.11
  
## Instalação (Linux):

```sh
$ sudo apt-get install -y libreoffice libreoffice-script-provider-python uno-libs3 python3-uno python3 
$ pip install unotools
$ pip install faker
```
## Execução (Linux):
 
### Iniciar o serviço do LibreOffice

```  
$ soffice --accept='socket,host=localhost,port=8100;urp;StarOffice.Service'
```  
ou
```  
$ soffice --accept='socket,host=localhost,port=8100;urp;StarOffice.Service' --headless
```
### Rodar o script
``` 
$ python script/find_replace.py -s localhost
``` 

## Arquivos

### Arquivos a serem processados : Pasta "input"
  - Formatos lidos pelo Writer ( libreoffice )

### Formato de Entrada - Substituições a serem realizadas : Pasta "subs":
  - Nome do Arquvo
  - Nome ( Nome completo a ser substituído )
  - Primeiro Nome
  - Sexo ( Masculino ou Feminino )

### Formato de Saída - Pasta "output"
  - Arquivos da pasta "input" processados
  - Arquvo resultado.csv ( contendo todas as substituiçes ) 
  - Arquivo de-para.csv ( contendo o de-para de todas as substituições )

## Futuras Implementações
  - Buscas variações do nome (Primeiro Nome, Nome sem acentos, Nome sem letras maiúsculas )
  - Ajustar formato de entradas
  - Refatorar
  
 
 
