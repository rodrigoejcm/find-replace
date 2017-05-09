# Find-Replace
Scritp utilizado para desidentificação de multiplos arquivos texto

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
OU
$ soffice --accept='socket,host=localhost,port=8100;urp;StarOffice.Service' --headless
```
### Rodar o script
``` 
$ python script/find_replace.py -s localhost
``` 
