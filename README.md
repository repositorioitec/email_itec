# Sistema de Gerenciamento de Emails (Protocolo IMAP) - Estilo Outlook

Sistema moderno de gerenciamento de correio eletrônico com suporte completo ao protocolo **IMAP** e **SMTP**, projetado com a interface e experiência de uso idênticas ao **Microsoft Outlook 365**.

Integrado com o servidor **ServidorHH / Roundcube** (`https://h58.servidorhh.com/roundcube/`).

---

## 🚀 Como Iniciar o Sistema

### Opção 1: Via executável direto (Windows)
Dê um duplo clique no arquivo:
```
iniciar_sistema.bat
```
O navegador abrirá automaticamente em [http://localhost:4000](http://localhost:4000).

### Opção 2: Via Terminal / Prompt de Comando
```bash
npm start
```

---

## 🌟 Principais Recursos

1. **Design e Interface Microsoft Outlook 365**:
   - Barra superior com pesquisa rápida por remetente, assunto ou conteúdo (`Ctrl+E`).
   - Barra de navegação lateral (Rail de aplicativos: Email, Calendário, Contatos, Tarefas).
   - Painel de pastas expansível: *Caixa de Entrada*, *Itens Enviados*, *Rascunhos*, *Lixo Eletrônico*, *Itens Excluídos*.
   - Contador de mensagens não lidas em tempo real.
   - Lista de emails com avatares, remetentes destacados, datas formatadas, sinalizadores e ações rápidas ao passar o mouse.
   - Painel de leitura e pré-visualização com suporte completo a HTML seguro e anexos.
   - Janela flutuante de **Nova Mensagem (Compose)** com editor de texto formatado, anexação de arquivos e envio via SMTP.

2. **Protocolo IMAP & SMTP**:
   - Conexão nativa e criptografada via SSL/TLS com o servidor `h58.servidorhh.com`.
   - **Porta IMAP de Entrada:** `993` (SSL habilitado).
   - **Porta SMTP de Saída:** `465` (SSL habilitado) ou `587`.
   - Gerenciamento de sinalizadores (marcar como lido/não lido, sinalizar com estrela).
   - Exclusão de mensagens e movimentação entre pastas.
   - Envio de emails com suporte a múltiplos destinatários, cópia (Cc), cópia oculta (Cco) e anexos.

3. **Painel de Configurações Interativo**:
   - Permite testar a conexão IMAP com 1 clique diretamente pela interface.
   - Salva e gerencia com segurança as credenciais do seu email.

---

## 📂 Estrutura do Projeto

- `server.js`: Servidor Express com rotas REST para operações IMAP e SMTP.
- `imap-service.js`: Motor de integração IMAP usando `imapflow` e `mailparser`.
- `smtp-service.js`: Motor de envio de emails usando `nodemailer`.
- `config.json`: Arquivo de configuração de portas e hosts do servidor de email.
- `public/`: Interface web cliente:
  - `index.html`: Estrutura do cliente Outlook 365.
  - `style.css`: Estilização visual no padrão Microsoft Fluent Design.
  - `app.js`: Lógica de gerenciamento de estado, chamadas de API e atalhos.


