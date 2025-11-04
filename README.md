# 📊 Controle de User Stories – Azure DevOps → Google Sheets

## 🎯 Objetivo

Automatizar a leitura das **User Stories** de todos os projetos do **Azure DevOps (CNC-TI)** e registrar, em uma planilha do Google Sheets, as principais informações de acompanhamento:

- Quem criou, validou, aceitou, resolveu e concluiu cada US.  
- As respectivas **datas** de cada etapa.  
- Status atual e link direto para o item no DevOps.  

Cada projeto é gravado em uma **aba separada** dentro da mesma planilha, chamada `Controle_US_Azure`.

---

## ⚙️ Arquitetura e Fluxo

1. **Azure DevOps API**  
   - Leitura dos projetos da organização via endpoint `/_apis/projects`.  
   - Consulta das User Stories via **WIQL**.  
   - Obtenção dos detalhes e histórico de cada Work Item.  

2. **Google Apps Script (GAS)**  
   - Código executado diretamente no Apps Script, conectado ao Google Drive.  
   - Cria ou atualiza a planilha automaticamente.  
   - Escreve os dados organizados em colunas padronizadas.

3. **Planilha no Google Drive**  
   - Uma aba por projeto (ex: SNCC, SEI, Representações, etc.).  
   - Cada execução limpa e reescreve os dados para manter tudo atualizado.

---

## 🧩 Estrutura do Código

| Seção | Descrição |
|-------|------------|
| **Configurações gerais** | Define variáveis de organização, token, pasta e API. |
| **Helpers genéricos** | Funções utilitárias: `extractIdentityName_`, `formatDate_`, `callAzureDevOps_`. |
| **Camada Azure DevOps** | Busca projetos, IDs e detalhes das US, e lê histórico de transições. |
| **Camada Google Sheets** | Cria ou abre a planilha e garante a estrutura de colunas. |
| **Orquestração** | Sincroniza cada projeto com sua aba na planilha. |
| **Funções de teste e debug** | Testes rápidos de token, WIQL e diagnósticos de itens específicos. |

---

## 📑 Campos Registrados

| Coluna | Origem / Lógica |
|--------|-----------------|
| **ID** | `System.Id` |
| **Link** | URL direta para a US no Azure DevOps |
| **Status** | `System.State` |
| **Criado por** | `System.CreatedBy` |
| **Data Criação** | `System.CreatedDate` formatado em `DD/MM/YYYY` |
| **Validado por** | Primeiro `revisedBy` que mudou o estado para `Ready` |
| **Data Validação** | Data da transição para `Ready` |
| **Aceito por** | Primeiro `revisedBy` que mudou o estado para `Active` |
| **Data Aceite** | Data da transição para `Active` |
| **Resolvido por** | Primeiro `revisedBy` que mudou o estado para `Resolved` ou campo `ResolvedBy` |
| **Data Resolvido** | Data da transição ou `ResolvedDate` |
| **Concluído por** | Primeiro `revisedBy` que mudou o estado para `Closed/Done` ou campo `ClosedBy` |
| **Data Concluído** | Data da transição ou `ClosedDate` |

---

## 🔑 Pré-requisitos

1. **Personal Access Token (PAT)** válido com permissão de leitura em *Work Items* e *Projects*.  
2. No Apps Script, criar uma **Script Property**:  
   - Nome: `AZURE_PAT`  
   - Valor: `<seu_token>`  
3. Garantir que exista a pasta no Google Drive com o ID definido em `FOLDER_ID`.

---

## 🚀 Funções Principais

| Função | Descrição |
|--------|------------|
| `syncAllProjects()` | Sincroniza **todos** os projetos da organização. |
| `syncProjectToSheet_(projectName)` | Sincroniza apenas um projeto específico. |
| `syncProjetoSNCC()` | Exemplo de teste individual (pode alterar o nome do projeto). |

---

## 🔄 Execução Automática (Opcional)

Para executar de forma agendada, adicione este trecho:

```javascript
function createSyncTrigger() {
  ScriptApp.newTrigger('syncAllProjects')
    .timeBased()
    .everyHours(6) // ou everyDays(1)
    .create();
}
