# RelogioDePonto
Programa em Python capaz de trabalhar como um relógio de ponto, exportando diretamente para XLSX.

- BREVE RESUMO

Essa é a minha mais nova criação em Python, um programa que faz tudo que é necessário para funcionar como um legítimo relógio de ponto.
Comentei praticamente tudo no meu código, então é possível compreender meu código sem dificuldade, mas ainda tenho várias revisões para fazer. 

Projetei inicialmente para que um janela do TKinter gerenciasse os serviços, mas ela está 80% funcional. Vou adicionar essa versão separadamente, mas não pretendo mais atualizar ela, visto que tenho um caminho melhor e mais prático.

Agora estou utilizando para gerenciar esse progama uma interface web, que ficou bem mais interativa e simples, assim qualquer usuário pode usar tranquilamente, visto que pensei que o uso seria para alguém responsável pelo RH.

- FUNCIONALIDADE

Essa interface em Python funciona como um simples CRUD, criando usuários no banco de dados, atualizando eles e deletando. Preferi por não deixar a opção de alterar nomes de funcionários, para não causar conflito. Após criado o programa gera um QR, que carrega a informação de ID do funcionário. Baseando-se apenas no ID, não há problema de haver mais um funcionário com o mesmo nome. Quando o programa "LerEscrever.py" verificar o QR, ele automaticamente faz todo o processo, verifica se há uma planilha existente, se é necessário criar, e por fim preenche toda a planilha. Quando ela estiver completamente preenchida, o programa guardará essa planilha cheia e iniciará uma nova. O programa "ReenvioXLSX.py" tem a função de reenviar planilhas, caso venha a acontecer algum problema na hora da gravação da planilha. É possível alterar informações necessárias, como IP ou senha do banco de dados pelo "Config.py". O processo é realizado de forma fácil se tudo for feito localmente, mas também é possível movimentar as planilhas para servidores remotos. Essa é uma função da qual desisti, devido ao trabalho excessivo e a grande quantidade de prevenção de erros, mas ela ainda estará disponível na versão do TKinter. Para utilizar a versão do TKinter, bastá executar o "Janela.py". Mais um ponto a se resaltar é que a só uma planilha especifica poderá conversar com esse programa, não qualquer arquivo .xlsx, que é o "Modelo.xlsx". É possível alterar para trabalhar com outras planilhas, você pode me contatar caso precise de ajuda com isso, mas essa planilha pronta é bem completa e realiza muitas funções automaticamente, então te recomendo fortemente a deixar da forma padrão.

- REQUISITOS

É possível trabalhar comprovadamente com Linux e Windows, sendo necessário pouco poder de processamento e baixo consumo de memória RAM. No Linux optei pro deixar os programas rodando o tempo todo por meio do SystemMD, sendo preciso apenas verificar permissões para gravar e editar planilhas nas pastas. No Windows ainda não encontrei uma solução para que o programa reinicie o tempo todo, caso feche por algum motivo. É necessário também uma câmera, para realizar a leitura do código QR. Também é preciso algumas bibliotecas que eu deixarei o método para instalação abaixo

Não é necessário excel ou qualquer outro leitor de arquivos .xlsx para utilizar esse programa, a não ser para visualizar as planilhas posteriormente.

Se você deseja utilizar a versão TKinter, é preciso que você acesse pasta "Add ao Diretorio", copie os dois executáveis que há dentro da pasta e cole eles dentro da pasta onde está o seu "python.exe". Normalmente esse é o caminho padrão na última versão -> C:\AppData\Local\Programs\Python\Python310

Me contate se necessário -> 1davydsantos@gmail.com
