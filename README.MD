# Automação de Envio de E-mails com Anexo SAP

Este projeto tem como objetivo automatizar o processo de envio de e-mails em massa com um arquivo extraído do sistema SAP. Através dessa automação, o sistema realiza uma conexão com o SAP para extrair um arquivo específico, salva o arquivo localmente e o envia por e-mail como anexo para uma lista de destinatários.

> **Nota:** Todas as informações empresariais sensíveis foram camufladas para fins de segurança e confidencialidade.

## Tecnologias Utilizadas

- **Python** (versão 3.8+)
- **SMTP** para envio de e-mails
- **SAP GUI Scripting** para interação com o SAP
- **Biblioteca Pywin32** para automação do SAP GUI (para Windows)
- **MIME** para criação de mensagens e-mails com anexos

## Pré-requisitos

1. **Python 3.8+** deve estar instalado.
2. **SAP GUI**: O cliente SAP GUI precisa estar instalado e configurado no ambiente.
3. **Configuração de e-mail SMTP**: Este exemplo utiliza o servidor SMTP do Gmail. Caso utilize outro provedor, ajuste as configurações de servidor e porta.
4. **Habilitação do SAP GUI Scripting**: Certifique-se de que o SAP GUI Scripting esteja ativado no seu ambiente SAP. Isso pode exigir permissões administrativas.

## Configurações Iniciais

1. **Instalação de Dependências**  
   Instale as dependências usando o comando abaixo:
   ```bash
   pip install pywin32
2. **Configuração do Servidor SMTP**
Para utilizar o Gmail como servidor SMTP, ative a autenticação de dois fatores e crie uma senha de aplicativo específica para essa automação.

## Estrutura do Projeto

1. main.py: Script principal que contém as funções de conexão com o SAP, extração de arquivo, e envio de e-mails em massa.
2. contatos.csv: Arquivo CSV com a lista de destinatários. Deve conter uma coluna para o nome e outra para o e-mail.
3. README.md: Documentação do projeto.

## Estrutura do Arquivo contatos.csv

O arquivo CSV com a lista de destinatários deve seguir o seguinte formato:
nome,email
Fulano,fulano@exemplo.com
Ciclano,ciclano@exemplo.com

## Como Funciona o Script
1. Conexão e Extração de Arquivo no SAP
O script usa SAP GUI Scripting para conectar-se ao SAP, acessar uma transação específica, executar o relatório e salvar o arquivo em um diretório local.

2. Envio de E-mails em Massa
O script então lê o arquivo contatos.csv e envia o e-mail para cada contato com o arquivo extraído do SAP como anexo. Cada e-mail é personalizado com o nome do destinatário.

## Observações Importantes
1. *Permissões no SAP*: Certifique-se de que o usuário SAP possui as permissões necessárias para acessar a transação e extrair o relatório.
1. *Testes de Conexão e Debugging*: SAP GUI Scripting pode apresentar diferenças dependendo do sistema. É importante realizar testes para garantir que o processo de extração e salvamento do arquivo esteja funcionando conforme o esperado.
3. *Confidencialidade*: As informações empresariais foram camufladas e devem ser substituídas pelas configurações reais no ambiente de produção.

## Autor e Licença
Desenvolvido por: Jean Lucas
Licença: MIT

