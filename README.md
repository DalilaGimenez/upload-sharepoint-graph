# Automatizador de Upload para SharePoint

Este projeto é um script de automação em Python projetado para fazer o upload de arquivos de uma pasta local para uma biblioteca de documentos do SharePoint Online. Ele organiza os arquivos em subpastas específicas com base em palavras-chave ou prefixos em seus nomes, envia relatórios de status por e-mail e mantém logs detalhados de todas as operações.

## Funcionalidades

-   **Autenticação Segura**: Utiliza o fluxo de credenciais de cliente OAuth 2.0 para se autenticar na API do Microsoft Graph de forma segura, sem expor senhas.
-   **Upload Inteligente**: Analisa os nomes dos arquivos e os move para subpastas pré-configuradas no SharePoint.
-   **Organização Automática**: Mapeia arquivos para pastas como `VENDAS`, `EQUIPE`, `CLIENTES`, etc., com base em regras personalizáveis no próprio código.
-   **Relatórios por E-mail**: Envia um e-mail de resumo após cada execução, notificando sobre sucessos e falhas. Em caso de erro crítico, um e-mail de alerta é disparado.
-   **Logging Completo**: Registra cada passo da execução em um arquivo de log com data e hora (`TaskSchedulerLog_YYYYMMDD_HHMMSS.txt`), facilitando a auditoria e a depuração.
-   **Portabilidade**: Projetado para rodar tanto como um script Python quanto como um executável independente (gerado com PyInstaller), ideal para tarefas agendadas.

## Pré-requisitos

-   Python 3.8 ou superior
-   Conta do Microsoft 365 com acesso de administrador para registrar uma aplicação no Azure AD.

## Instalação

1.  **Clone o repositório** (ou baixe os arquivos para um diretório):
    ```bash
    git clone <url-do-seu-repositorio>
    cd <nome-do-diretorio>
    ```

2.  **Crie e ative um ambiente virtual** (recomendado):
    ```bash
    # Windows
    python -m venv venv
    .\venv\Scripts\activate

    # macOS / Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Instale as dependências** a partir do arquivo `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```

## Configuração

Antes de executar o script, você precisa configurar as variáveis de ambiente.

1.  **Registre uma Aplicação no Azure Active Directory**:
    -   Vá para o portal do Azure > Azure Active Directory > Registros de aplicativo.
    -   Crie um novo registro.
    -   Anote o `ID do Aplicativo (cliente)` (CLIENT_ID) e o `ID do Diretório (locatário)` (TENANT_ID).
    -   Em "Certificados e segredos", crie um novo segredo do cliente e anote o **Valor** (CLIENT_SECRET).
    -   Em "Permissões de API", adicione as permissões `Sites.ReadWrite.All` e `Mail.Send` (do tipo "Aplicativo") e conceda o consentimento do administrador.

2.  **Crie o arquivo `.env`**:
    -   Renomeie ou copie o arquivo `.env.example` para `.env`.
    -   Preencha as variáveis com as informações obtidas no passo anterior e com os detalhes do seu ambiente.

## Uso

Para executar o script, basta rodar o arquivo `main.py` a partir do seu terminal (com o ambiente virtual ativado):

```bash
python main.py
```

O script irá:
1.  Autenticar-se na API da Microsoft.
2.  Verificar a pasta de origem (`SOURCE_FOLDER`).
3.  Fazer o upload dos arquivos para as pastas corretas no SharePoint.
4.  Enviar um e-mail de relatório.
5.  Salvar um arquivo de log no mesmo diretório do script.

### Agendamento de Tarefa (Exemplo para Windows)

Você pode usar o Agendador de Tarefas do Windows para executar este script automaticamente.

-   **Ação**: Iniciar um programa
-   **Programa/script**: `C:\caminho\completo\para\seu\venv\Scripts\python.exe`
-   **Adicione argumentos (opcional)**: `C:\caminho\completo\para\seu\projeto\main.py`
-   **Iniciar em (opcional)**: `C:\caminho\completo\para\seu\projeto\`

## Estrutura do Projeto

```
.
├── helpers/
│   ├── __init__.py
│   ├── auth.py         # Lógica de autenticação
│   ├── sharepoint.py   # Funções de interação com o SharePoint
│   └── email.py        # Função para envio de e-mail
├── .env                # Arquivo de configuração (não versionado)
├── .env.example        # Exemplo de arquivo de configuração
├── main.py             # Ponto de entrada da aplicação
├── requirements.txt    # Lista de dependências Python
└── README.md           # Este arquivo
```