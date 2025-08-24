Automatizador de Sincronização de Planilhas Microsoft Forms
Este é um script em Python que automatiza o processo de abrir planilhas do Excel Online vinculadas ao Microsoft Forms para forçar a sincronização de dados. Ele utiliza o Selenium para controlar um navegador em modo headless (sem interface gráfica), realizando login, navegando até as planilhas especificadas e aguardando a confirmação de que os dados foram sincronizados.

O Problema Resolvido
O Microsoft Forms permite que as respostas de formulários sejam salvas automaticamente em uma planilha do Excel Online. No entanto, em alguns cenários, a sincronização de novas respostas só ocorre quando a planilha é aberta manualmente por um usuário. Este script resolve esse problema ao simular a ação de um usuário, garantindo que as planilhas estejam sempre atualizadas sem intervenção manual.

Principais Funcionalidades
Login Automático: Faz login em uma conta Microsoft de forma segura usando credenciais armazenadas em um arquivo de ambiente.

Operação em Modo Headless: Roda em segundo plano, sem abrir janelas de navegador visíveis, ideal para execução em servidores ou agendamento de tarefas (como Cron Jobs ou Agendador de Tarefas do Windows).

Verificação Inteligente de Sincronização: Não depende de esperas fixas. O script procura ativamente pela mensagem "Sincronizado com Forms" na página, inclusive dentro de iframes, tornando o processo mais rápido e confiável.

Gerenciamento Automático do ChromeDriver: Utiliza webdriver-manager para baixar e gerenciar a versão correta do driver do Chrome, eliminando a necessidade de atualizações manuais.

Configuração Flexível de URLs: Permite adicionar múltiplas planilhas, organizadas por categorias, de forma simples e clara.

Captura de Tela para Depuração: Em caso de falha no login ou em outros pontos críticos, o script salva uma captura de tela para facilitar a identificação do problema.

Estrutura Robusta: Utiliza blocos try...finally para garantir que o navegador seja sempre encerrado corretamente, mesmo em caso de erros.

Pré-requisitos
Antes de começar, garanta que você tenha os seguintes itens instalados:

Python 3.7+

Google Chrome instalado na máquina onde o script será executado.

Uma conta Microsoft com acesso às planilhas que você deseja sincronizar.

Instalação
Siga os passos abaixo para configurar o ambiente do projeto.

1 . Clone o repositório:

git clone https://github.com/RaphaelClemente/Automatizador-de-Planilhas-do-Formulario-Microsoft

2 . Crie e ative um ambiente virtual (recomendado):

No Windows: python -m venv venv .\venv\Scripts\activate

3 . Instale as dependências:
Crie um arquivo requirements.txt com o seguinte conteúdo:

selenium
python-dotenv
webdriver-manager

Em seguida, instale as bibliotecas com o pip:

pip install -r requirements.txt

Configuração
A configuração do script é feita através de duas partes principais:

1. Credenciais de Acesso (.env)
Para manter suas credenciais seguras, o script utiliza um arquivo .env.

1.Crie um arquivo chamado .env na raiz do projeto.

2.Adicione suas credenciais da Microsoft neste arquivo:
EMAIL_MS="seu_email@exemplo.com"
SENHA_MS="sua_senha_super_secreta"

Importante: Adicione o arquivo .env ao seu .gitignore para evitar que suas credenciais sejam enviadas para o repositório Git.


2. Lista de Planilhas (no script)
Abra o arquivo Python principal e edite o dicionário urls para adicionar as planilhas que você deseja processar. Mantenha a estrutura de ("Nome da Planilha", "URL_completa").

# Lista de URLs para verificar
urls = {
    "Vendas": [
        ("Relatório Diário", "https://office.live.com/start/Excel.aspx?om=1&id=..."),
        ("Feedback de Clientes", "https://link-para-sua-outra-planilha")
    ],
    "Recursos Humanos": [
        ("Pesquisa de Clima", "https://link-para-planilha-de-rh")
    ],
    "Outros": []
}

Como Usar
Com o ambiente configurado e o script devidamente ajustado, basta executá-lo a partir do seu terminal:

python nome_do_seu_script.py

Configurando o navegador para rodar em modo headless...
Navegador iniciado em segundo plano.
Login realizado com sucesso!

O script iniciará o processo, exibindo o progresso no terminal:

Categoria: Vendas
   Acessando: Relatório Diário
    Encontrados 1 iframes. Verificando dentro deles...
    Sincronização detectada dentro do iframe #0!
    Status: OK
   Acessando: Feedback de Clientes
    Sincronização detectada na página principal!
    Status: OK

Todas as planilhas foram processadas!

Fechando o navegador...



Como Funciona
1.Carregamento: As credenciais são carregadas do arquivo .env.

2.Inicialização do Navegador: Uma instância do Google Chrome é iniciada em modo headless com configurações otimizadas para automação.

3.Login: O script navega para a página de login da Microsoft, insere o e-mail e a senha, e lida com o prompt de "Manter conectado".

4.Processamento em Loop: O script itera sobre cada categoria e planilha definida no dicionário urls.

5.Acesso e Verificação: Para cada planilha, ele acessa a URL e chama a função esperar_sincronizacao.

6.Espera Inteligente: A função de espera procura ativamente pela mensagem de sucesso "Sincronizado com Forms". Ela é capaz de encontrar o elemento tanto na página principal quanto dentro de elementos iframe, que são comuns em aplicações web complexas como o Office Online.

7.Finalização: Após processar todas as URLs, o script encerra a sessão do navegador de forma segura.
