#-----------------------------------------------------------------------

#Todos os pacotes necessários

from Pacotes import *
from Color   import *
from Config  import *

#Importando conexão com banco de dados

from MySQL import *

#-----------------------------------------------------------------------

#Minimziar/Ocultar Terminal

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

#-----------------------------------------------------------------------

#Gerador de Log
#Inserir diretório onde os arquivo de log deve ficar
#Se ele já existir, irá escrever nele
#Se não existir, ele irá gerar um novo e irá escrever nele

def Estragui():

    Log = open("log_file.txt", "a")

    IDFunc = "1"

    #-----------------------------------------------------------------------

    #Verificação de conexão com o banco de dados, servidor HTTP e Internet
    #Se qualquer um dos três falhar, o programa ira falhar
    #Você pode verificar a situação no log_file.txt
    #HTTP Desativado. Utilizando shuttil por ser mais simples
    #SQL e DNS Desativado, não necessário já que o server MySQL será local

    #PingSQL  = str(ping(IPSQL))
    #PingHTTP = str(ping("192.168.15.128"))
    #PingDNS  = str(ping(IPDNS))

    #if PingSQL == "None" or PingSQL == "False":
        #print("Conexão com o banco de dados: "+Falha)
        #print("Verifique a conexão e tente novamente")
        #Log.write("\n LE.PY -> "+str(datetime.now())+" - FALHA! - Sem acesso ao banco de dados.")
        #Estragui()
        
    #elif PingSQL != "None" or PingSQL != "False":
        #print("Conexão com o banco de dados: "+Sucesso)
        #Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - Acesso ao banco de dados checado.")
        
    #if PingHTTP == "None" or PingHTTP == "False":
        #print("Conexão com o servidor HTTP: Fail")
        #print("Verifique a conexão e tente novamente")
        #Log.write("\n"+str(datetime.now())+" - FALHA! - Sem acesso ao servidor HTTP.")
        
    #elif PingHTTP != "None" or PingHTTP != "False":
        #print("Conexão com o servidor HTTP: OK!")
        #Log.write("\n"+str(datetime.now())+" - SUCESSO! - Acesso ao servidor HTTP checado.")

    #if PingDNS == "None" or PingDNS == "False":
        #print("Conexão com a internet: "+Falha)
        #print("Verifique a conexão e tente novamente")
        #Log.write("\n LE.PY -> "+str(datetime.now())+" - FALHA! - Sem acesso a internet.")
        #Estragui()
        
    #elif PingDNS != "None" or PingDNS != "False":
        #print("Conexão com a internet: "+Sucesso)
        #Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - Acesso a internet checado.")

    #-----------------------------------------------------------------------

    #Virado é uma váriavel pra ajustar o IF de quem bater ponto depois da 00:00

    Virado = False

    #-----------------------------------------------------------------------

    #Inicio Captura QR
    #Pode ser necessário trocar o index da variável cap para outro, afim de encontrar a câmera
    #Mas o padrão é 0

    cap      = cv2.VideoCapture(0)
    detector = cv2.QRCodeDetector()

    while True:
        _,img        = cap.read()
        data, bbox, _= detector.detectAndDecode(img)

        if data:
            IDFunc   = data
            Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - Leitura de QR com sucesso. ID identificado: "+IDFunc)
            break

        cv2.imshow("QRCODEscanner", img)   

        #Linha para parar leitura com uma tecla. Não necessária
        if cv2.waitKey(1) == ord("q"):
            Log.write("\n LE.PY -> "+str(datetime.now())+" - AVISO! - Leitura de QR cancelada.")
            #os.system("start cmd /c python LerEscrever.py")
            #exit()
            break

    #-----------------------------------------------------------------------

    #Obter ID pelo QR e selecionar nome da tabela
    #Código QR pode ser obtido no "CadastroFuncionario.py"

    mycursor = mydb.cursor()
    mycursor.execute("SELECT nome FROM "+NomeTabela+" WHERE id = "+IDFunc+"")
    myresult = mycursor.fetchall()

    #-----------------------------------------------------------------------

    #Pequena verificação, caso a ID não seja encontrada no Banco

    if myresult == []:
        print(Falha+" - ID "+IDFunc+" não encontrada no banco de dados")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - FALHA! - ID "+IDFunc+" nao encontrada no banco de dados.")
        Estragui()

    else:
        print(Sucesso+" - ID "+IDFunc+" encontrada no banco de dados.")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - ID "+IDFunc+" encontrada no banco de dados.")

    #NomeFuncA = Nome do funcionário com caracteres
    #Retirar = Caracteres para retirar do nome do funcionário
    #Logo após, função para retirar caracteres do nome do funcionário
    #Nome do Funcionário = NomeFunc

    NomeFuncA = myresult[0]
    Retirar   = "(',)"
    NomeFunc  = "".join(char for char in NomeFuncA if char not in Retirar)

    #-----------------------------------------------------------------------

    #Após, CV2 finaliza
    
    cap.release()
    cv2.destroyAllWindows()

    #-----------------------------------------------------------------------

    #Para rodar server HTTP, python -m http.server
    #O diretório inicial do HTTP será na pasta onde o script foi rodado, porta 8000
    #Sempre duplicar "\" no código, não necessário com "/"

    #LinkExcel = Local da planilha para download. Pode usar arquivos locais e desativar função
    #Open realiza a função de salvar o arquivo, IDFunc - NomeFunc.xlsx
    #Exemplo: "10 - Davyd da Silva Santos.xlsx"

    #LinkExcel  = "http://192.168.15.128:8000/CartaoPonto/"+str(IDFunc)+" - "+NomeFunc+".xlsx"
    #Requisicao = requests.get(LinkExcel, allow_redirects = True)
    #open("C:\\Users\\PC_2019\\Documents\\"+str(IDFunc)+" - "+NomeFunc+".xlsx", "wb").write(Requisicao.content)

    #-----------------------------------------------------------------------

    #Código acima está desativado
    #Metodo mais simples será aplicado usando shutil.copy()
    #PlanilhaCopiada será para onde é enviada a cópia que sera editada
    #LocalPraCopia será de onde é copiado a planilha  

    PlanilhaCopiada = DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx"
    LocalPraCopia   = DiretorioServer+str(IDFunc)+" - "+NomeFunc+".xlsx"

    #Planilha é o local da planilha. Sempre duplicar "\" no código, não necessário com "/"
    #Capa é o nome da planilha interna
    #DataInicioComHora é celula que diz a data de inicial utilizada como base para planilha interna "MODELO"
    #DataInicioSTR é igual a de cima, porem em formato de string
    #DataInicio é igual a de cima, porem sendo apenas data, sem a hora junto
    #DataPy é a data atual obtida pelo Python
    #DiferencaDias é a diferença de dias entre o dia atual (DataPy) e a da de inicio da planilha (DataInicio)
    #LinhaPraEscrita é onde sera escrito no cartão de ponto quando se bater o cartão, referente a data atual
    #HoraMinutoAtual é a hora e o minuto atuais
    #HoraAtual é apenas a hora atual, multiplicada por 100, para poder concaternar com o MinutoAtual
    #MinutoAtual é apenas o minuto atual
    #HoraPraEscrita é a hora e minuto atuais concatenados que serão escritos na celula selecionada

    try:

        shutil.copy(LocalPraCopia, PlanilhaCopiada)
        Planilha          = load_workbook(filename = DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")
        Capa              = Planilha.worksheets[2]
        DataInicioComHora = str(Capa['C4'].value)
        DataInicioSTR     = DataInicioComHora[0:10]
        DataInicio        = datetime.strptime(DataInicioSTR, '%Y-%m-%d').date()

        DataPy            = date.today()

        DiferencaDias     = abs((DataInicio - DataPy).days)
        LinhaPraEscrita   = 7 + DiferencaDias

        HoraeMinutoAtual  = datetime.now()
        HoraAtual         = HoraeMinutoAtual.hour * 100
        MinutoAtual       = HoraeMinutoAtual.minute
        HoraPraEscrita    = HoraAtual + MinutoAtual

        #Agora, a cada quinze dias a planilha se renovara
        #Se utilizara uma vazia como base
        #As antigas serão movidas para uma pasta e terão "antiga" no seu nome
        #Será necessária no minimo uma manutenção quinzenal
        #O restante do código permanece o mesmo de antes, escrevendo o nome e a data inicial na planilha

        if LinhaPraEscrita > 36:

            print(Aviso+" - Planilha já chegou no limite e será trocada por uma nova")
            Log.write("\n LE.PY -> "+str(datetime.now())+" - AVISO! - Planilha sendo trocada por uma nova.")
            Planilha.save(DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")
            shutil.copy(DiretorioServer+str(IDFunc)+" - "+NomeFunc+".xlsx", DiretorioCheia+str(IDFunc)+" - "+NomeFunc+" - "+str(DataInicio)+".xlsx")
            Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - Planilha cheia armazenada com sucesso.")
            shutil.copy(DiretorioModelo, DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")

            Planilha          = load_workbook(filename = DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")
            Capa              = Planilha.worksheets[2]
            Modelo            = Planilha.worksheets[4]
            Capa["C4"]        = date.today()
            Modelo["D4"]      = NomeFunc
            DataInicioComHora = str(Capa['C4'].value)
            DataInicioSTR     = DataInicioComHora[0:10]
            DataInicio        = datetime.strptime(DataInicioSTR, '%Y-%m-%d').date()

            DataPy            = date.today()

            DiferencaDias     = abs((DataInicio - DataPy).days)
            LinhaPraEscrita   = 7 + DiferencaDias

            HoraeMinutoAtual  = datetime.now()
            HoraAtual         = HoraeMinutoAtual.hour * 100
            MinutoAtual       = HoraeMinutoAtual.minute
            HoraPraEscrita    = HoraAtual + MinutoAtual

    except:
        print(Falha+" - Não existe planilha para esta funcionário. Será gerada um nova planilha.")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - FALHA! - Nao existe planilha para esta funcionario. Sera gerada um nova planilha.")
        Existencia = False

        shutil.copy(DiretorioModelo, DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")

        Planilha          = load_workbook(filename = DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")
        Capa              = Planilha.worksheets[2]
        Modelo            = Planilha.worksheets[4]
        Capa["C4"]        = date.today()
        Modelo["D4"]      = NomeFunc
        DataInicioComHora = str(Capa['C4'].value)
        DataInicioSTR     = DataInicioComHora[0:10]
        DataInicio        = datetime.strptime(DataInicioSTR, '%Y-%m-%d').date()

        DataPy            = date.today()

        DiferencaDias     = abs((DataInicio - DataPy).days)
        LinhaPraEscrita   = 7 + DiferencaDias

        HoraeMinutoAtual  = datetime.now()
        HoraAtual         = HoraeMinutoAtual.hour * 100
        MinutoAtual       = HoraeMinutoAtual.minute
        HoraPraEscrita    = HoraAtual + MinutoAtual

    #-----------------------------------------------------------------------

    #Se o linha para escrever passar da linha 37, não haverá celula pronta pra escrita
    #Portanto, será necessário atualizar a data de início
    #Essa verificação é pra que isso não aconteça, será necessária uma atualização constante das planilhas
    #Para que não excedam os limites da planilhas

    #if LinhaPraEscrita > 37:
        #print("O cartão de ponto está cheio")
        #Log.write("\n"+str(datetime.now())+" - FALHA! - Tentativa de bater ponto pelo funcionario "+IDFunc+" - "+NomeFunc+", porem o cartao esta cheio.")
        #Log.close()
        #exit()

    #-----------------------------------------------------------------------

    #Código acima está desativado

    #-----------------------------------------------------------------------

    #Neste if se verifica se o dia atual é o dia inicial da planilha
    #Se for, não há como verificar se todos os pontos foram batidos no dia anterior
    #Pois não há planilha acima
    #False = Não há consulta do dia anterior
    #True  = Há consulta do dia anterior

    if LinhaPraEscrita == 7:
        ProcuraDiaAnterior = bool(False)

    elif LinhaPraEscrita > 7:
        ProcuraDiaAnterior = bool(True)

    #-----------------------------------------------------------------------

    #Modelo é o nome da planilha interna

    Modelo = Planilha.worksheets[4]

    #Respectivamente: Entrada, Entrada Almoço, Saída Almoço, Saída

    Entrada = str(Modelo["D"+str(LinhaPraEscrita)].value)
    EAlmoco = str(Modelo["E"+str(LinhaPraEscrita)].value)
    SAlmoco = str(Modelo["F"+str(LinhaPraEscrita)].value)
    Saida   = str(Modelo["G"+str(LinhaPraEscrita)].value)

    #-----------------------------------------------------------------------

    #Neste if, LinhaPraProcurar será a linha do dia anterior
    #EntradaDiaAnterior é o valor da celula da Entrada do dia anterior
    #EAlmocoDiaAnterior é o valor da celula da Entrada do Almoço do dia anterior
    #SAlmocoDiaAnterior é o valor da celula da Saída do Almoço do dia anterior
    #SaidaDiaAnterior é o valor da celula da Saída do dia anterior

    if ProcuraDiaAnterior == True:
        LinhaPraProcurar   = LinhaPraEscrita - 1
        EntradaDiaAnterior = str(Modelo["D"+str(LinhaPraProcurar)].value)
        EAlmocoDiaAnterior = str(Modelo["E"+str(LinhaPraProcurar)].value)
        SAlmocoDiaAnterior = str(Modelo["F"+str(LinhaPraProcurar)].value)
        SaidaDiaAnterior   = str(Modelo["G"+str(LinhaPraProcurar)].value)

    #Se as entradas e saídas do dia anteirior forem vazia, fica claro que o dia anterior foi folga ou falta
    #Após isso, ira escrever na linha do dia atual, como entrada

    #-----------------------------------------------------------------------

    #Verificação de batida do último ponto
    #Aguardar cinco minutos ao menos para o mesmo funcionário bater o ponto novamente

    mycursor.execute("SELECT id FROM "+NomeTabela+" WHERE ultimo_ponto < NOW() - INTERVAL 10 MINUTE and id = "+IDFunc)
    myresult = mycursor.fetchall()

    if myresult == []:
        print(Aviso+" - Ultimo ponto batido a menos de dez minutos. Aguarde para bater o ponto novamente.")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - AVISO! - O funcionario "+IDFunc+" - "+NomeFunc+" tentou bater o ponto novamente em menos de dez minutos.")
        Estragui()

    else:
        print(Sucesso+" - Ultimo ponto batido a mais de dez minutos. Liberado.")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - Verificacao de tempo limite para o funcionario "+IDFunc+" - "+NomeFunc+" realizada. Liberado.")


    #-----------------------------------------------------------------------

    if ProcuraDiaAnterior == True:

        if EntradaDiaAnterior == "None" and EAlmocoDiaAnterior == "None" and SAlmocoDiaAnterior == "None" and SaidaDiaAnterior == "None":
            print(Aviso+" - Você não teve expediente no dia anterior. Caso não seja folga ou falta, informe o RH.")
            #Desativado por não fazer sentido
            #Escrevimento = "D"+str(LinhaPraEscrita)
            #Modelo[Escrevimento] = 1234
            Log.write("\n LE.PY -> "+str(datetime.now())+" - AVISO! - O funcionario "+IDFunc+" - "+NomeFunc+" nao teve expediente ontem.")

    #Se a ultima celula do dia anterior estiver vazia
    #É necessário escrever nela, em caso de se bater ponto após 00:00
    #Para que o programa não entenda que se deve escrever no dia atual, após 00:00
    #Escrevimento se define como local onde se deve escrever

        elif EntradaDiaAnterior != "None" and EAlmocoDiaAnterior != "None" and SAlmocoDiaAnterior != "None" and SaidaDiaAnterior == "None":
            print(Sucesso+" - Você bateu o ponto na saida de ontem.")
            Escrevimento = "G"+str(LinhaPraProcurar)
            Modelo[Escrevimento] = HoraPraEscrita
            Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - O funcionario "+IDFunc+" - "+NomeFunc+" bateu o ponto da saida de ontem.")
            
            mycursor.execute("UPDATE "+NomeTabela+" SET ultimo_ponto = NOW()  WHERE id = "+IDFunc)
            mydb.commit()
        
            Log.write("\n LE.PY -> "+str(datetime.now())+" - AVISO! O funcionario "+IDFunc+" - "+NomeFunc+" bateu o ponto apos 00:00.")
            Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! horario do funcionario "+IDFunc+" - "+NomeFunc+" atualizado no banco.")
            Virado = True

    #Caso contrário, não alterar nada e escrever no dia atual

        elif SaidaDiaAnterior != "None":
            print(Aviso+" - Ultimo dia completamente preenchido.")

    #Fim desta verificação

    #-----------------------------------------------------------------------

    #Aqui se da continuidade na verificação, referente ao dia atual
    #Se acabar de preencher a saída 00:00, não se deve escrever na entrada
    #Ao escrever, verifica se já foi dado entrada. Se não foi, escrever na entrada
    #Se já tiver dado entrada, verifica se foi para o almoço. Se não foi, escrever na entrada almoço 
    #Se já tiver dado entrada no almoço, verifica se saiu do almoço. Se não foi, escrever na saida almoço
    #Se já tiver dado saida no almoço, verifica se já saiu da empresa. Se não, escrever na saida

    if Virado == True and Entrada == "None":
        Virado = False
    
    elif Entrada == "None":
        print(Sucesso+" - Você bateu o ponto de entrada.")
        Escrevimento = "D"+str(LinhaPraEscrita)
        Modelo[Escrevimento] = HoraPraEscrita
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - O funcionario "+IDFunc+" - "+NomeFunc+" bateu o ponto de entrada.")

        mycursor.execute("UPDATE "+NomeTabela+" SET ultimo_ponto = NOW() WHERE id = "+IDFunc)
        mydb.commit()
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! horario do funcionario "+IDFunc+" - "+NomeFunc+" atualizado no banco.")
        
    elif Entrada != "None" and EAlmoco == "None":
        print(Sucesso+" - Você saiu para o almoço.")
        Escrevimento = "E"+str(LinhaPraEscrita)
        Modelo[Escrevimento] = HoraPraEscrita
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - O funcionario "+IDFunc+" - "+NomeFunc+" saiu para o almoco.")

        mycursor.execute("UPDATE "+NomeTabela+" SET ultimo_ponto = NOW() WHERE id = "+IDFunc)
        mydb.commit()
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! horario do funcionario "+IDFunc+" - "+NomeFunc+" atualizado no banco.")

    elif EAlmoco != "None" and SAlmoco == "None":
        print(Sucesso+" - Você voltou do almoço.")
        Escrevimento = "F"+str(LinhaPraEscrita)
        Modelo[Escrevimento] = HoraPraEscrita
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - O funcionario "+IDFunc+" - "+NomeFunc+" voltou do almoco.")

        mycursor.execute("UPDATE "+NomeTabela+" SET ultimo_ponto = NOW() WHERE id = "+IDFunc)
        mydb.commit()
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! horario do funcionario "+IDFunc+" - "+NomeFunc+" atualizado no banco.")


    elif SAlmoco != "None" and Saida == "None":
        print(Sucesso+" - Você bateu o ponto de saída.")
        Escrevimento = "G"+str(LinhaPraEscrita)
        Modelo[Escrevimento] = HoraPraEscrita
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - O funcionario "+IDFunc+" - "+NomeFunc+" bateu o ponto de saida.")

        mycursor.execute("UPDATE "+NomeTabela+" SET ultimo_ponto = NOW() WHERE id = "+IDFunc)
        mydb.commit()
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! horario do funcionario "+IDFunc+" - "+NomeFunc+" atualizado no banco.")


    #Se tentar escrever no dia atual e tudo já estiver preenchidp
    #Não se deve reescrever. O ideal será verificar com o RH a situação da planilha
    #E ver se há uma falha na programação
    #Deve ser incomum, mas pode ser que haja duas pessoas com o mesmo QR
    #Ou o leitor possa ler incorretamente um QR, apesar de improvável

    elif Entrada != "None" and EAlmoco != "None" and SAlmoco != "None" and Saida != "None":
        print(Aviso+" - Saída de hoje já preenchida anteriormente com sucesso. Caso não esteja correto, informe o RH.")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - AVISO! - O funcionario "+IDFunc+" - "+NomeFunc+" tentou bater novamente o ponto de saida atual.")

    #Fim desta verificação

    #-----------------------------------------------------------------------

    #Salvar planilha

    Planilha.save(DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx")

    #-----------------------------------------------------------------------

    #Devolvendo planilha para o servidor
    #PlanilhaShutil é o local onde está a planilha atualizada
    #ReenvioShutil é onde ficará armazenado os arquivos que não conseguiram ser enviados
    #Se a planilha estiver em uso no server, não é possível regravar o arquivo
    #Então o programa, ReenvioXLSX.py, irá aguardar ser possível e tentar novamente
    #ServerShutil é o local onde deve ser salvo a planilha

    PlanilhaShutil = DiretorioUtilizando+str(IDFunc)+" - "+NomeFunc+".xlsx"
    ServerShutil   = DiretorioServer+str(IDFunc)+" - "+NomeFunc+".xlsx"
    ReenvioShutil  = DiretorioReenviar+str(IDFunc)+" - "+NomeFunc+".xlsx"

    try:
        MeuArquivo = open(DiretorioServer+str(IDFunc)+" - "+NomeFunc+".xlsx", "r+")
        shutil.move(PlanilhaShutil,ServerShutil)
        print(Sucesso+" - Planilha devolvida ao destino com sucesso")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - SUCESSO! - Arquivo "+IDFunc+" - "+NomeFunc+".xlsx enviado com sucesso!")
        Log.close()
    except IOError:
        shutil.move(PlanilhaShutil,ReenvioShutil)
        print(Aviso+" - O arquivo "+str(IDFunc)+" - "+NomeFunc+".xlsx não foi enviado ao destino. Tentando novamente.")
        Log.write("\n LE.PY -> "+str(datetime.now())+" - FALHA! - Arquivo "+IDFunc+" - "+NomeFunc+".xlsx nao foi enviado. Tentando novamente.")
        Log.close()

    #-----------------------------------------------------------------------

    #Reinicia o programa

while True:
    Estragui()

#Código abaixo desabilitado pois fazia o programa abrir novamente sem fechar o anterior
#Utilizando essa gambiara com While, o PID permanece sempre o mesmo, o que da um total controle do programa
#os.system("start cmd /c python LerEscrever.py")
#exit()