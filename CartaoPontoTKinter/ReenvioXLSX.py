#-----------------------------------------------------------------------

#Todos os pacotes necessários

from Pacotes import *
from Color   import *
from Config  import *

#-----------------------------------------------------------------------

#Minimziar/Ocultar Terminal

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

#-----------------------------------------------------------------------

#NArquivo é a variável que dita o arquivo da lista

def Fuioroto():

    NArquivo = 0

#-----------------------------------------------------------------------

#Inicio do laço

    while True:

    #Seleciona diretório e verifica o que há dentro

        def ListarArquivos(Caminho = None):
            ArquivosNaPasta = [arq for arq in listdir(Caminho)]
            return (ArquivosNaPasta)

        if __name__ == '__main__':
            Arquivos = ListarArquivos(Caminho = DiretorioReenviar)

    #Se for vazia, não trabalho. Reinicia após 10 segundos

            if Arquivos == []:
                print(Aviso+" - Sem arquivos para reenviar. Leitura em 10 segundos.")
                break
            
    #Se houver arquivo, se definem varíaveis de local e destino
    #Respectivamente "Reenvio" e "Server"
        
            elif Arquivos != []:
                ReenvioShutil  = DiretorioReenviar+Arquivos[NArquivo]
                ServerShutil   = DiretorioServer+Arquivos[NArquivo]

    #Se tenta enviar o arquivo, array 0
    #Se .move for bem sucedido, o arquivo é enviado com sucesso
    #Se lê o diretório novamente
    #Verifica quantos documentos tem dentro do diretório
    #Se o diretório estiver vazio, fim do laço. Reinicia após 10 segundos
                
                try:
                    print(Aviso+" - Tentativa de enviar arquivo "+Arquivos[NArquivo])
                    shutil.move(ReenvioShutil,ServerShutil)
                    print(Sucesso+" - Arquivo enviado com sucesso!")
                    Arquivos = ListarArquivos(Caminho = DiretorioReenviar)
                    QuantArq = len(Arquivos)
                    if Arquivos == []:
                        print(Sucesso+" - Todos os arquivos foram enviados com sucesso. Leitura em 30 segundos!")
                        break
                    
    #Se .move não for bem sucedida
    #Se lê o diretório novamente
    #Verifica quantos documentos tem dentro do diretório
    #Avisa que .move não foi bem sucedido
    #Aguarda 2 segundos
                    
                except IOError:
                    Arquivos = ListarArquivos(Caminho = DiretorioReenviar)
                    QuantArq = len(Arquivos)
                    print(Falha+" - Arquivo sendo usado por outro usuário. Tentando enviar novamente o arquivo "+Arquivos[NArquivo]+".")
                    time.sleep(30)

    #NArquivo passa a ser o array 1
                    
                NArquivo = NArquivo + 1

    #Se o array foi maior que quantidade de arquivos, então ele não existe
    #E o array torna a ser 0
    #Caso contrário, vai se somando até o array ser todos os arquivos
                
                if NArquivo >= QuantArq:
                    NArquivo = 0
                    
    ##Se o diretório estiver vazio, seja por qual motivo, novamente fim do laço. Reinicia após 30 segundos
                    
                if Arquivos == []:
                    print(Sucesso+" - Todos os arquivos foram enviados com sucesso. Leitura em 30 segundos!")
                    break

    #-----------------------------------------------------------------------

    time.sleep(30)

while True:
    Fuioroto()

#Sistema abaixo não utilizado
#Necessário ser o mesmo programa por conta do PID

#os.system("start cmd /c python ReenvioXLSX.py")
#exit()

#-----------------------------------------------------------------------

#Feito para rodar infinitamente, com pausa para baixo consumo de memória