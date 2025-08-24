# 📊 Automatizador de Sincronização de Planilhas Microsoft Forms

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)  
[![Selenium](https://img.shields.io/badge/Selenium-Automation-brightgreen.svg)](https://www.selenium.dev/)  
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)  

---

## 📌 Sobre o Projeto
Este é um script em **Python** que automatiza o processo de abrir planilhas do **Excel Online** vinculadas ao **Microsoft Forms** para forçar a sincronização de dados.  

Ele utiliza o **Selenium** para controlar um navegador em modo **headless** (sem interface gráfica), realizando login, navegando até as planilhas especificadas e aguardando a confirmação de que os dados foram sincronizados.

---

## ❓ O Problema Resolvido
O Microsoft Forms permite que as respostas de formulários sejam salvas automaticamente em uma planilha do Excel Online.  
No entanto, em alguns cenários, a sincronização de novas respostas só ocorre quando a planilha é aberta manualmente.  

➡️ Este script resolve esse problema ao simular a ação de um usuário, garantindo que as planilhas estejam sempre atualizadas sem intervenção manual.

---

## ⚙️ Principais Funcionalidades
- 🔑 **Login Automático** com credenciais seguras (.env).  
- 🕶️ **Execução em Modo Headless** (ideal para servidores e agendadores).  
- ⚡ **Verificação Inteligente de Sincronização** (inclusive dentro de iframes).  
- 🔄 **Gerenciamento Automático do ChromeDriver** via `webdriver-manager`.  
- 📂 **Configuração Flexível de URLs** por categorias.  
- 📸 **Captura de Tela para Depuração** em caso de falhas.  
- 🛡️ **Estrutura Robusta** com tratamento de erros e encerramento seguro do navegador.  

---

## 🔧 Pré-requisitos
Antes de começar, você precisa ter instalado:  

- Python **3.7+**  
- Google Chrome  
- Conta Microsoft com acesso às planilhas desejadas  

---

## 🚀 Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/RaphaelClemente/Automatizador-de-Planilhas-do-Formulario-Microsoft

2.Crie e ative um ambiente virtual (recomendado):

Windows:
python -m venv venv
.\venv\Scripts\activate

3.Instale as dependências:
Crie um arquivo requirements.txt com:
selenium
python-dotenv
webdriver-manager

Em seguida, rode:
pip install -r requirements.txt

⚙️ Configuração
1. Credenciais de Acesso (.env)

Crie um arquivo chamado .env na raiz do projeto com suas credenciais Microsoft:
EMAIL_MS="seu_email@exemplo.com"
SENHA_MS="sua_senha_super_secreta"

2. Lista de Planilhas (no script)

No arquivo Python principal, edite o dicionário urls para adicionar as planilhas:
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

▶️ Como Usar

Com tudo configurado, execute no terminal:
python nome_do_seu_script.py

Saída esperada no terminal:
Configurando o navegador para rodar em modo headless...
Navegador iniciado em segundo plano.
Login realizado com sucesso!

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


🔍 Como Funciona

1.Carregamento: As credenciais são lidas do arquivo .env.

2.Inicialização do Navegador: Chrome é iniciado em modo headless.

3.Login: O script acessa a página de login da Microsoft e autentica o usuário.

4.Processamento em Loop: Itera sobre as categorias e planilhas no dicionário urls.

5.Acesso e Verificação: Abre cada planilha e procura pela mensagem "Sincronizado com Forms".

6.Espera Inteligente: Suporta verificação dentro de iframes.

7.Finalização: Encerra o navegador de forma segura após processar todas as URLs.

📜 Licença

Este projeto está licenciado sob a licença MIT.
Você é livre para usar, modificar e distribuir este projeto, desde que mantenha a atribuição ao autor original.

👨‍💻 Desenvolvido por Raphael Clemente
   
