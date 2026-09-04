# Regras do Projeto Kanban

## 1. Versionamento Obrigatório
- **Toda e qualquer alteração** no código exige incrementar o número de versão (`APP_VERSION`) no `Kanban.html` (ex: `v1.0.43` → `v1.0.44`).
- O cache do Service Worker em `sw.js` deve ser incrementado no mesmo commit (ex: `kanban-v10-cache` → `kanban-v11-cache`).

## 2. Consistência entre Kanban e Apps Scripts
- Sempre que houver ajustes em **regras de validação, expressões regulares (regex), formatação de links ou tratamento de chaves (keys/tickets)** no `Kanban.html`:
  - **TODOS os scripts da pasta `Apps Scripts/`** devem ser auditados e atualizados obrigatoriamente para seguirem exatamente as mesmas regras e comportamentos.
- Nenhuma alteração em lógica de parsing (ex: regex de Jira, regex de Zendesk, hiperlinks) pode ser feita apenas no frontend sem refletir em todos os scripts do Google Apps Script correspondentes.

### Onde fica cada script (um arquivo por projeto Apps Script)
| Arquivo no repo | Projeto Apps Script | Contém |
|---|---|---|
| `Apps Scripts/Comite.gs` | planilha **Acompanhamento Comitê** | `doGet` (Web App do Kanban) + painel/menus + sync Comitê→Fábrica |
| `Apps Scripts/Tickets Encaminhado para Fábrica/Código.gs` | planilha **Tickets Encaminhados Para a Fábrica** | `onEdit` (links) + sync Apoios→Comitê |

- **Um arquivo por projeto, de propósito.** Os `.gs` de um projeto compartilham o mesmo escopo global — separar em vários arquivos nunca isolou nada, só multiplicava o que colar ao atualizar. Ao criar função nova, siga o prefixo da seção: `kb*` (API do Kanban), `pn*` (painel/menus), `ce*` (sync). Nomes já usados por gatilho/implantação **não podem mudar**: `doGet`, `onOpen`, `processarEdicaoComite`, `atualizarPainel`, `criarTriggerAuto`, `excluirTicketComite`.
- Alterou o `Comite.gs` em algo que o Kanban consome? Publicar de novo: **Implantar > Gerenciar implantações > editar > Versão: Nova versão** (a URL `/exec` não muda e os gatilhos não precisam ser reinstalados).
