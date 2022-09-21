#-----------------------------------------------------------------------

#Todos os pacotes necessários

from Pacotes import *
from Color   import *
from Config  import *

#-----------------------------------------------------------------------

#Minimziar/Ocultar Terminal

ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

#-----------------------------------------------------------------------

VerificaLE = False
VerificaRE = False

GlobalPIDLE = 0
GlobalPIDRE = 0

StatusLE = "Iniciar"
CorLE = "#FF9D9D"

StatusRE = "Iniciar"
CorRE = "#FF9D9D"

def RunConfig():
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    os.system("start notepad Config.py")

def RunLog():
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    os.system("start notepad log_file.txt")

def LLerEscrever():
    os.system("start cmd /c LerEscrever LerEscrever.py")

    LerEscreverAcao["text"] = "Parar"
    LerEscreverAcao["command"] = DLerEscrever 
    LerEscrever["bg"] = "#9DFF9D"
    VerificaLE = True
    Console.configure(state = "normal")
    Console.insert(INSERT, "JL.PY -> "+str(datetime.now())+" - SUCESSO! - Bater o ponto iniciado.\n")
    Console.configure(state = "disabled")    
    Console.yview(END)

    time.sleep(3)

    for proc in psutil.process_iter():
        if proc.name() == "LerEscrever.exe":
            global GlobalPIDLE
            PIDLerEscrever = proc.pid
            GlobalPIDLE = PIDLerEscrever
            LerEscreverPID["text"] = PIDLerEscrever            


def DLerEscrever():

    LerEscreverAcao["text"] = "Iniciar"
    LerEscreverAcao["command"] = LLerEscrever
    LerEscrever["bg"] = "#FF9D9D"
    VeriricaLE = False
    Console.configure(state = "normal")
    Console.insert(INSERT, "JL.PY -> "+str(datetime.now())+" - PARADO! - Bater o ponto parado.\n")
    Console.configure(state = "disabled")    
    Console.yview(END)

    for proc in psutil.process_iter():
        if proc.pid == GlobalPIDLE:
            proc.kill()
            LerEscreverPID["text"] = "0"
            break

def LReenviar():
    os.system("start cmd /c ReenviarXLSX ReenvioXLSX.py")

    ReenviarAcao["text"] = "Parar"
    ReenviarAcao["command"] = DReenviar 
    Reenviar["bg"] = "#9DFF9D"
    VerificaRE = True
    Console.configure(state = "normal")
    Console.insert(INSERT, "JL.PY -> "+str(datetime.now())+" - SUCESSO! - Reenviar planilhas inciado.\n")
    Console.configure(state = "disabled")    
    Console.yview(END)

    time.sleep(3)

    for proc in psutil.process_iter():
        if proc.name() == "ReenviarXLSX.exe":
            global GlobalPIDRE
            PIDReenviar = proc.pid
            GlobalPIDRE = PIDReenviar
            ReenviarPID["text"] = PIDReenviar

def DReenviar():

    ReenviarAcao["text"] = "Iniciar"
    ReenviarAcao["command"] = LReenviar
    Reenviar["bg"] = "#FF9D9D"
    VerificaRE = True
    Console.configure(state = "normal")
    Console.insert(INSERT, "JL.PY -> "+str(datetime.now())+" - PARADO! - Reenviar planilhas parado.\n")
    Console.configure(state = "disabled")    
    Console.yview(END)


    for proc in psutil.process_iter():
        if proc.pid == GlobalPIDRE:
            proc.kill()
            ReenviarPID["text"] = "0"
            break

def RunCadastro():
    os.system("start cmd /c python CadastroFuncionario.py")
    Console.configure(state = "normal")
    Console.insert(INSERT, "JL.PY -> "+str(datetime.now())+" - SUCESSO! - Cadastro iniciado.\n")
    Console.configure(state = "disabled")    
    Console.yview(END)

def RunSair():

    for proc in psutil.process_iter():
        if GlobalPIDRE == 0:
            break
        elif proc.pid == GlobalPIDRE:
            proc.kill()
            break

    for proc in psutil.process_iter():
        if GlobalPIDLE == 0:
            break
        elif proc.pid == GlobalPIDLE:
            proc.kill()
            break
        
    while True:
        for proc in psutil.process_iter():
            if proc.name() == "python.exe":
                proc.kill()
                break

def disable_event():
    pass

if VerificaLE == False:
    ComandoLE = LLerEscrever

elif VerificaLE == True:
    ComandoLE = DLerEscrever

if VerificaRE == False:
    ComandoRE = LReenviar
    
elif VerificaRE == True:
    ComandoRE = DReenviar


Janela = Tk()

Janela.protocol("WM_DELETE_WINDOW", disable_event)
Janela.iconbitmap("Icones\\Relogio.ico")
Janela.geometry("700x300")
Janela.resizable(False, False)

Janela.title("Relógio Ponto v1.0b")

Botoes = Frame(Janela)
Botoes.pack()

TextoCOne = Label(Botoes, text = "Serviços") 
TextoCOne.grid(column = 0, row = 0)

TextoCTwo = Label(Botoes, text = "PID")
TextoCTwo.grid(column = 1, row = 0)

TextoCThree = Label(Botoes, text = "Açao")
TextoCThree.grid(column = 2, row = 0)

LerEscrever = Label(Botoes, text = "Bater o ponto", bg = CorLE, height = 1, width = 16)
LerEscrever.grid(column = 0, row = 1)

LerEscreverPID = Label(Botoes, text = GlobalPIDLE, height = 1, width = 9)
LerEscreverPID.grid(column = 1, row = 1)

LerEscreverAcao = Button(Botoes, text = StatusLE, height = 1, width = 10, command = ComandoLE)
LerEscreverAcao.grid(column = 2, row = 1)

Reenviar = Label(Botoes, text = "Envio de planilha", bg = CorRE, height = 1, width = 16)
Reenviar.grid(column = 0, row = 2)

ReenviarPID = Label(Botoes, text = GlobalPIDRE, height = 1, width = 9)
ReenviarPID.grid(column = 1, row = 2)

ReenviarAcao = Button(Botoes, text = StatusRE, height = 1, width = 10, command = ComandoRE)
ReenviarAcao.grid(column = 2, row = 2)

Cadastro = Label(Botoes, text = "Cadastrar usuário", height = 1, width = 16)
Cadastro.grid(column = 0, row = 3)

CadastroPID = Label(Botoes, text = "", height = 1, width = 9)
CadastroPID.grid(column = 1, row = 3)

CadastroAcao = Button(Botoes, text = "Cadastrar", height = 1, width = 10, command = RunCadastro)
CadastroAcao.grid(column = 2, row = 3)

ColunaEspaco = Label(Botoes, text = "", height = 1, width = 4)
ColunaEspaco.grid(column = 3, row = 0)

ConfigIco = PhotoImage(file = "Icones\\Config.png")
ConfigAcao = Button(Botoes, text = "Configurar", height = 20, width = 90, image = ConfigIco, compound = LEFT, command = RunConfig)
ConfigAcao.grid(column = 4, row = 1)

LogIco = PhotoImage(file = "Icones\\Lupa.png")
LogAcao = Button(Botoes, text = "Log", height = 20, width = 90, image = LogIco, compound = LEFT, command = RunLog)
LogAcao.grid(column = 4, row = 2)

SairIco = PhotoImage(file = "Icones\\Sair.png")
SairAcao = Button(Botoes, text = "Sair", height = 20, width = 90, image = SairIco, compound = LEFT, command = RunSair)
SairAcao.grid(column = 4, row = 3)

Console = ScrolledText(Janela, height = 10, width = 80)
Console.configure(state = "disabled")
Console.pack(pady = 30)

Janela.mainloop()