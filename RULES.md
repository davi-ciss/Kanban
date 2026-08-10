# Regras do Projeto Kanban

## 1. Versionamento Obrigatório
- **Toda e qualquer alteração** no código exige incrementar o número de versão (`APP_VERSION`) no `Kanban.html` (ex: `v1.0.43` → `v1.0.44`).
- O cache do Service Worker em `sw.js` deve ser incrementado no mesmo commit (ex: `kanban-v10-cache` → `kanban-v11-cache`).

## 2. Consistência entre Kanban e Apps Scripts
- Sempre que houver ajustes em **regras de validação, expressões regulares (regex), formatação de links ou tratamento de chaves (keys/tickets)** no `Kanban.html`:
  - **TODOS os scripts da pasta `Apps Scripts/`** (incluindo `KanbanAPI.gs.txt`, `Scripts.gs.txt`, `Código.gs.txt` e subpastas) devem ser auditados e atualizados obrigatoriamente para seguirem exatamente as mesmas regras e comportamentos.
- Nenhuma alteração em lógica de parsing (ex: regex de Jira, regex de Zendesk, hiperlinks) pode ser feita apenas no frontend sem refletir em todos os scripts do Google Apps Script correspondentes.
