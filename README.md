
# 🧩 Controle de User Stories – Azure DevOps → Google Sheets

## 📘 Visão Geral
Este projeto integra **User Stories do Azure DevOps** com o **Google Sheets**, automatizando o controle e acompanhamento de status de cada US em múltiplos projetos.  
Além de sincronizar as informações, o sistema também gera **logs de execução** e **envia notificações por e-mail** quando há alterações de status entre execuções.

---

## 🚀 Objetivos do Script

- Ler **User Stories** dos projetos configurados no Azure DevOps.  
- Calcular responsáveis e datas de cada etapa:
  - **Validado**, **Aceito**, **Resolvido** e **Concluído**.  
- Gravar todas as informações em uma **planilha única no Google Drive**, com **uma aba por projeto**.  
- Registrar um histórico de execução (aba `LOG_EXECUCAO`).  
- Enviar **e-mail automático** quando forem detectadas alterações de status nas US.

---

## ⚙️ Requisitos

### 1. Azure DevOps
- Criar um **Personal Access Token (PAT)** com permissão de leitura em *Work Items*.  
- Armazenar o token em:
  ```
  Propriedades do Script → Propriedades do Projeto → AZURE_PAT
  ```

### 2. Google Drive
- Criar uma **pasta dedicada** (ex: “Azure”) no seu Google Drive.  
- Obter o **ID da pasta** (presente na URL) e preencher no código:
  ```javascript
  const FOLDER_ID = 'SEU_ID_DA_PASTA_AQUI';
  ```

- Se for usar uma **conta técnica**, ela precisa ter **permissão de Editor** nessa pasta.

### 3. Planilha
- O script cria automaticamente (na primeira execução) a planilha:
  ```
  Controle_US_Azure
  ```
  dentro da pasta configurada em `FOLDER_ID`.  
- Se ela já existir, será atualizada — uma aba por projeto (ex: `SNCC`, `SEI`, etc.).

---

## 🧠 Estrutura do Script

### Principais funções:

| Função | Descrição |
|--------|------------|
| `listProjects_()` | Lista todos os projetos ativos na organização do Azure DevOps. |
| `fetchUserStoryIds_(projectName)` | Retorna os IDs das User Stories via WIQL. |
| `fetchUserStoriesDetails_(projectName, ids)` | Busca detalhes das US em lotes. |
| `getFirstTransitionInfo_(...)` | Determina quem e quando mudou uma US para um determinado estado. |
| `getOrCreateSpreadsheet_()` | Cria ou abre a planilha `Controle_US_Azure` dentro da pasta configurada. |
| `ensureProjectSheet_()` | Garante que a aba do projeto existe e escreve o cabeçalho. |
| `syncProjectToSheet_(projectName)` | Sincroniza um projeto específico (lê, calcula e grava na aba). |
| `syncAllProjects()` | Sincroniza todos os projetos da organização. |
| `notifyUserStoryChanges_(currentItems)` | Detecta mudanças de status e envia e-mail automático. |

---

## 📨 Envio de E-mails Automáticos

### Configurações principais

```javascript
const EMAIL_NOTIFICATIONS_ENABLED = true;  // define se o envio está ativo
const EMAIL_FROM_NAME = 'Christian Moura dos Santos';
const EMAIL_REPLY_TO = 'christian_7c@cnc.org.br';
const EMAIL_FROM_ADDRESS = 'contactconsultservices@gmail.com';  // conta técnica
const EMAIL_TO = 'sistemas@cnc.org.br';  // destinatário padrão
```

### Conteúdo do e-mail

- O e-mail inclui:
  - Data da execução
  - Tabela com: Projeto, ID, Status anterior, Status atual, **Data da ação** e link direto da US
  - Link para a planilha completa
- As cores seguem a paleta corporativa CNC:
  - Azul escuro `#004d73`
  - Branco `#ffffff`
  - Cinza claro `#f5f9fc`

Exemplo de assunto:
```
Atualização de User Stories – Azure DevOps
```

---

## 🗂 Estrutura da Planilha

Cada aba (um projeto) contém:

| Coluna | Descrição |
|--------|------------|
| ID | Identificador da User Story |
| Link | Link direto no Azure DevOps |
| Status | Estado atual |
| Criado por | Autor original |
| Data Criação | Data de criação |
| Validado por / Data Validação | Quem validou e quando |
| Aceito por / Data Aceite | Responsável por ativar a US |
| Resolvido por / Data Resolvido | Quem marcou como resolvido |
| Concluído por / Data Concluído | Quem finalizou e data de fechamento |

Além disso, existe a aba `LOG_EXECUCAO`:
| Projeto | ID | Status | Link |

Essa aba é usada para detectar alterações entre execuções.

---

## 🔁 Gatilhos (Triggers)

- Configure um gatilho **baseado em tempo** (ex: diário às 07:00) para a função:
  ```javascript
  syncAllProjects
  ```
- **Importante:**  
  O gatilho precisa estar **criado na conta técnica** (ex: `contactconsultservices@gmail.com`)  
  — é ela quem precisa ter acesso ao Drive e à planilha.

---

## 🧩 Fluxo Resumido de Execução

1. **syncAllProjects()**
   - Lista projetos do Azure DevOps.
   - Para cada um:
     - Busca US e escreve na planilha.
     - Retorna um snapshot com ID, status e datas.

2. **notifyUserStoryChanges_()**
   - Compara o snapshot atual com o `LOG_EXECUCAO`.
   - Identifica US novas ou com mudança de status.
   - Atualiza o `LOG_EXECUCAO`.
   - Se houver mudanças, monta o e-mail e envia.

3. **buildEmailBody_()**
   - Monta o HTML com cores CNC, tabela e link da planilha.

---

## 🧩 Testes e Diagnóstico

| Função | Uso |
|--------|-----|
| `testAzureToken()` | Valida se o PAT e o acesso ao Azure estão corretos. |
| `testListUserStories()` | Retorna amostra de IDs de User Stories. |
| `debugConclusao_5523()` | Debug detalhado de transição de status (exemplo real de US). |
| `testNotifySingleProject()` | Sincroniza um projeto e dispara e-mail de teste com as mudanças detectadas. |

---

## ⚠️ Permissões e Cuidados

- Se o erro for:
  ```
  Exception: Access denied: DriveApp.
  ```
  então a conta técnica **não tem acesso à pasta do Drive** (`FOLDER_ID`).

  ➜ Solução:  
  No Drive, compartilhe a pasta com a conta técnica como **Editor**.

- O script só pode enviar e-mails pela conta que o executa.
  Se quiser usar uma conta técnica para envio, ela precisa rodar o script.

- Se o `EMAIL_NOTIFICATIONS_ENABLED` estiver `false`, o log será atualizado mas nenhum e-mail será enviado.

---

## 🧾 Histórico de Evolução

| Data | Alteração |
|------|------------|
| 2025-10 | Versão inicial da integração Azure → Sheets |
| 2025-11 | Adicionado controle de LOG_EXECUCAO e e-mail automático |
| 2025-11 | Implementadas cores CNC e link direto da planilha |
| 2025-11 | Correção de permissão de DriveApp e suporte a conta técnica |
