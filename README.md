# Relogio de Ponto
Programa em Python capaz de trabalhar como um relógio de ponto, exportando diretamente para XLSX.

### ● BREVE RESUMO

Essa é a minha mais nova criação em Python, um programa que faz tudo que é necessário para funcionar como um legítimo relógio de ponto.
Comentei praticamente tudo no meu código, então é possível compreender sem dificuldade, mas ainda tenho várias revisões para fazer. 

Durante o projeto eu mudei algumas ideias e fiz duas versões, uma utilizando o TKinter e outra uma interface web. Deixarei as duas versões disponíveis mas não pretendo trazer atualizações para a versão TKinter.

Projetei inicialmente para que um janela do TKinter gerenciasse os serviços, ela está 80% funcional. Você pode encontrar ela na pasta [CartaoPontoTKinter](CartaoPontoTKinter). Você pode iniciar todo o processo da verão TKinter executando [Janela](CartaoPontoTKinter/Janela.py)

Agora estou utilizando uma interface web para gerenciar esse progama. Dessa forma ficou bem mais interativa e simples, assim qualquer usuário pode usar tranquilamente.

### ●	FUNCIONALIDADE

Essa interface em Python funciona como um simples CRUD, criando usuários no banco de dados, atualizando eles e deletando. Preferi por não deixar a opção de alterar nomes de funcionários, para não causar conflito. Após criado o programa gera um QR, que carrega a informação de ID do funcionário. Baseando-se apenas no ID, não há problema de haver mais um funcionário com o mesmo nome. Quando o programa [LerEscrever](CartaoPonto/LerEscrever.py) verificar o QR, ele automaticamente faz todo o processo, verifica se há uma planilha existente, se é necessário criar, e por fim preenche toda a planilha. Quando ela estiver completamente preenchida, o programa guardará essa planilha cheia e iniciará uma nova. O programa [ReenvioXLSX](CartaoPonto/ReenvioXLSX.py) tem a função de reenviar planilhas, caso venha a acontecer algum problema na hora da gravação da planilha. O processo é realizado de forma fácil se tudo for feito localmente, mas também é possível movimentar as planilhas para servidores remotos. Essa é uma função da qual desisti, devido ao trabalho excessivo e a grande quantidade de prevenção de erros, mas é possível implementa-la novamente sem grandes transtornos. 

Um ponto a se resaltar é que só uma planilha [Modelo.xlsx](CartaoPonto/Modelo.xlsx) poderá conversar com esse programa, não qualquer arquivo .xlsx. É possível alterar para trabalhar com outras planilhas, você pode me contatar caso precise de ajuda com isso, mas essa planilha pronta é bem completa e realiza muitas funções referente a horas trabalhadas automaticamente, então te recomendo fortemente a utilizar a planilha padrão.

Você pode configurar diretórios e variáveis para conexão com o banco em [Config](CartaoPonto/Config.py)

Você pode encontrar as planilhas em [Planilhas](CartaoPonto/Planilhas)

Você pode encontrar planilhas cheias em [CheiaXLSX](CartaoPonto/CheiaXLSX)

Você pode encontrar o QR do funcionário em [QRCode](CartaoPonto/QRCode)

### ● REQUISITOS

● Será necessário instalar o [Python](https://www.python.org/downloads/); 

● Será necessário instalar o [MySQL](https://dev.mysql.com/downloads/installer/), eu estou utilizando a versão do link em meus testes;

● É possível trabalhar comprovadamente com Linux e Windows, sendo necessário pouco poder de processamento e pouca memória RAM. No Linux optei pro deixar os programas rodando o tempo todo por meio do [systemd](https://www.freedesktop.org/software/systemd/man/systemd.service.html), sendo preciso apenas verificar permissões para gravar e editar planilhas nas pastas. No Windows ainda não encontrei uma solução para que o programa reinicie o tempo todo, caso feche por algum motivo. Ensinarei como fiz o systemd posteriomente; 

● É necessário também uma câmera, para realizar a leitura do código QR;

● Não é necessário excel ou qualquer outro leitor de arquivos .xlsx para utilizar esse programa, a não ser para visualizar as planilhas posteriormente;

● Se você deseja utilizar a versão TKinter, é preciso que você acesse pasta [Add ao Diretorio](CartaoPontoTKinter/Add ao Diretorio), copie os dois executáveis que há dentro da pasta e cole eles dentro da pasta onde está o seu *python.exe*. Normalmente esse é o caminho padrão na última versão -> *C:\AppData\Local\Programs\Python\Python310*;

● Também é preciso algumas bibliotecas que eu deixarei abaixo:

### ● BIBLIOTECAS NECESSÁRIAS

Você pode ver para que uso cada biblioteca em [Pacotes](CartaoPonto/Pacotes.py). Utilize esse [bat](CartaoPonto/InstallPack.bat) para Windows caso queira agilizar e instalar todas as bibliotecas de uma vez. 

● [OpenCV](https://pypi.org/project/opencv-python/)

● [MySQL Connector Python](https://pypi.org/project/mysql-connector-python/)

● [OpenPyXL](https://pypi.org/project/openpyxl/)

● [PyPNG](https://pypi.org/project/pypng/)

● [PSutil](https://pypi.org/project/psutil/)

● [PyQRCode](https://pypi.org/project/PyQRCode/)

● [PyTest Shutil](https://pypi.org/project/pytest-shutil/)

● [Python Time](https://pypi.org/project/python-time/)

● [Requests](https://pypi.org/project/requests/)

● Já faz parte do python, mas instale mesmo assim -> **pip3 install datetime**

### ● REGRAS E DICAS

Pensei em várias possibilidades que poderiam acontecer quando esse programa rodar, por exemplo, se alguém bater o ponto depois da meia-noite, ou em caso de folga, férias, ou bater o ponto novamente depois de já ter completado o dia, então o programa deve escrever tudo na planilha de forma correta. Algo a se observar é que você não pode fazer bater o ponto um atrás do outro, é necessário esperar 10 minutos. Fiz isso porque a leitura é bem rápida e seria um problema se o funcionário batesse todos os pontos do dia de uma vez.

1. Se utilizar a versão TKinter, ao iniciar o "Bater o ponto", não se esqueça de iniciar "Envio de planilha". É possível funcionar sem ele, mas ele previne erros.

2. O arquivo [Query](CartaoPonto/Query.txt) já contém o query para você criar o banco e a tabela exatamente como já está configurado.

3. Se você quer utilizar o programa sem modificá-lo, recomendo que cole a pasta *"CartaoPonto"* diretamente na raiz *"C:"*.

4. É possível usar um servidor externo para armazenamento das planilhas apenas mudando o texto dá variável do "DiretorioServer" dentro de [Config](CartaoPonto/Config.py).

5. É possível armazenar as planilhas no banco de dados, mas por padrão preferi armazená-las localmente.

6. É possível fazer o sistema funcionar online, interligando lojas ou para cumprir outra função.

7. Não utilizei a versão TKinter no Linux então não sei se é estável e 100% funcional.

8. Se está tendo algum problema na execução do sistema, verifique o [Log File](CartaoPonto/log_file.txt) ou rode o script isoladamente em uma IDE com depuração ou no IDLE do Python.

**_Me contate se necessário -> 1davydsantos@gmail.com_**
