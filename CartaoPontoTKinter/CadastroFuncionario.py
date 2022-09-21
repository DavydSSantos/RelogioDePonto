#-----------------------------------------------------------------------

#Todos os pacotes necessários

from Pacotes import *
from Color   import *
from Config  import *

#-----------------------------------------------------------------------

#Importando conexão com banco de dados

from MySQL import *

#-----------------------------------------------------------------------


#Gerador de Log
#Inserir diretório onde os arquivo de log deve ficar
#Se ele já existir, irá escrever nele
#Se não existir, ele irá gerar um novo e irá escrever nele

Log = open("log_file.txt", "a")

#-----------------------------------------------------------------------

#NomeFunc é o nome do funcionário que irá para o banco

NomeFunc = input("Digite o nome do funcionário(a): ")

#Se o nome for vazio, irá pedir para digitar novamente
#Loop até que o nome não seja vazio

while NomeFunc == "":
    NomeFunc = input("Digite o nome do funcionário(a): ")

Confirma = input("O nome '" + NomeFunc + "' está correto? S/N ")

#Se você confirmar ou negar que o nome esteja correto
#Você entrará em um novo laço
#Se você confirmar, fim do laço
#Se você negar, redigita o nome do funcionário é entra em um laço igual ao anterior
#Se você digitar algo diferente de S ou N, pedirá que confirme novamente

while True:

    if Confirma == "S" or  Confirma == "s":
        break

    elif Confirma == "N" or Confirma == "n":
        NomeFunc = input("Redigite o nome do funcionário(a): ")
        
        while NomeFunc == "":
            NomeFunc = input("Digite o nome do funcionário(a): ")
            
        Confirma = input("O nome '" + NomeFunc + "' está correto? S/N ")

    elif Confirma != "S" or Confirma != "s" or Confirma != "N" or Confirma != "n":
        print("Entrada incorreta")
        Confirma = input("O nome '" + NomeFunc + "' está correto? S/N ")

Log.write("\n CE.PY -> "+str(datetime.now())+" - AVISO! - Nome do funcionario: "+NomeFunc+".")

#-----------------------------------------------------------------------

#Conexão com banco
#Inserindo no banco as informações

mycursor = mydb.cursor()

mycursor.execute("INSERT INTO "+NomeTabela+" (nome) VALUES ('"+NomeFunc+"')")
mydb.commit()

#Verifica ultimo ID que foi inserido no banco
#Variável IDFunc

IDFunc = mycursor.lastrowid
print(Sucesso+" - Funcionário inserido com Sucesso!")
print(Aviso+" - O ID do(a) funcionário(a) '"+NomeFunc+"', é "+str(IDFunc)+".")

Log.write("\n CE.PY -> "+str(datetime.now())+" - SUCESSO! - Funcionario "+NomeFunc+" adicionado com sucesso.")
Log.write("\n CE.PY -> "+str(datetime.now())+" - AVISO! - ID do funcionario "+NomeFunc+": "+str(IDFunc)+".")

#-----------------------------------------------------------------------

#Gera um QRCode com o ID da pessoa codificado
#Salvo no diretório com nome, por exemplo, "10 - Davyd Santos.png"

QR = pyqrcode.create(IDFunc)

#É possível alterar o tamanho do QR caso queira mudando o "scale" abaixo:

QR.png(DiretorioQRCode+str(IDFunc)+" - "+NomeFunc+".png", scale = 6)

print(Sucesso+" - QRCode gerado com Sucesso!")
print(Aviso+" - Salvo em -> "+DiretorioQRCode+str(IDFunc)+" - "+NomeFunc+".png")

Log.write("\n CE.PY -> "+str(datetime.now())+" - SUCESSO! - Codigo QR do funcionario "+NomeFunc+" criado com sucesso.")
Log.write("\n CE.PY -> "+str(datetime.now())+" - AVISO! - QR armazenado na pasta QRCODE.")

wait = input("Pressione Enter para sair.")

Log.close()

quit()

#NÃO SE ESQUEÇA
#SE PRECISAR REDIFINIR ID
#ALTER TABLE nome_table AUTO_INCREMENT = 1