# Kanban — v1

App de organização **100% local**, estilo Trello, conectado à **sua aba** da planilha de acompanhamento **só quando você quiser**. Multiusuário: cada pessoa põe o nome da própria aba em ⚙ → *Seu nome (aba da planilha)*.

## Como usar (offline, sem configurar nada)

1. Duplo-clique em **`Kanban.html`** → abre no navegador.
2. Digite em "**+ Adicionar tarefa**" no rodapé de qualquer coluna e aperte **Enter**.
3. **Arraste** cards entre colunas; **clique** num card para editar (título, detalhe, ticket, cliente, prioridade).
4. Tudo é salvo no navegador desta máquina (localStorage). Nada vai pra internet.

Atalhos: **Ctrl+F** busca · **Esc** fecha · clique no nome da coluna para renomear.

Colunas padrão: *A Fazer · Em andamento · Aguardando · Concluído* (edite/adicione/exclua à vontade).

> Backup: engrenagem ⚙ → **Exportar backup (.json)**. Como a pasta é versionada (`v1`, `v2`…), guarde o `.json` ou copie a pasta antes de grandes mudanças.

## Backup / migrar de navegador ou máquina

Os dados ficam no **localStorage** deste navegador. Para backup ou pra levar pra outro navegador/máquina:

- ⚙ → **⤴ Exportar backup (.json)** gera um arquivo com tudo (cores, ordem, filtros, aparência e cards).
- No outro navegador/máquina: ⚙ → **⤵ Importar backup** e selecione esse `.json`.

Guarde o `.json` no Google Drive de vez em quando. (Os cards sincronizados com o Sheets você também recupera com **Puxar chamados**.)

## Instalar como atalho-app no Windows (opcional, sem servidor)

```
powershell -NoProfile -ExecutionPolicy Bypass -File "Instalar-Kanban.ps1"
```
Cria **Kanban** no Desktop e no Menu Iniciar (pesquise "Kanban"), com ícone próprio, abrindo em **janela própria** (Edge/Chrome `--app`, sem abas). Não é o "Instalar" nativo de PWA (esse exige servidor), mas na prática fica igual a um app. Pra desinstalar, apague os `.lnk`.

## Conectar com a sua aba (opcional)

O card só vira linha na planilha quando você ativa a sincronização. Para habilitar:

### 1. Publicar o endpoint (uma vez, ~1 min)
1. Abra a planilha de acompanhamento → menu **Extensões > Apps Script**.
2. Use o arquivo **`Apoios N3/KanbanAPI.gs`** (cole/atualize no projeto da planilha) e salve.
3. **Implantar > Nova implantação** → engrenagem → **Aplicativo da Web**.
   - *Executar como:* **Eu**
   - *Quem pode acessar:* **Qualquer pessoa**
4. **Implantar** → autorize → copie a **URL do app da Web** (termina em `/exec`).

### 2. Apontar o Kanban
1. No Kanban, engrenagem ⚙ → cole a URL em **Endpoint** e preencha **Seu nome (aba da planilha)** com o nome exato da sua aba.
2. **Testar conexão** → deve mostrar `✓ Conectado · aba "<seu nome>" (N linhas)`.
3. **⤓ Puxar chamados da aba** importa as linhas como cards (Status → coluna correspondente).

### Depois de conectado
- **Vincular** um card → cria/atualiza a linha (Ticket, Assunto→Assunto, Cliente, Status, Observação).
- Arrastar um card **vinculado** para outra coluna → atualiza o **Status** e o **Alteração de Status** na planilha automaticamente.
- Cards não vinculados continuam só no Kanban.

## Mapeamento de colunas

O `Code.gs` acha as colunas pelo **nome do cabeçalho** (linha 1), em qualquer ordem, ignorando acentos:

| Card            | Cabeçalho na aba |
|-----------------|------------------|
| Ticket          | Ticket / Chamado |
| Título          | Assunto / Título |
| Cliente         | Cliente |
| Coluna do board | Status |
| Detalhe         | Observação |
| (automático)    | Alteração de Status |

Se a aba usar outros nomes, ajuste `COLMAP` no topo do `Code.gs`.

## Limites / notas
- A comunicação usa **JSONP** (tag `<script>`), o que evita problemas de CORS rodando do `file://`. Por isso o endpoint precisa ser "Qualquer pessoa".
- Payload via URL: ok para cards normais; observações gigantes (>1–2 mil caracteres) podem falhar — encurte se necessário.
- Sem internet, vínculo fica indisponível, mas o board funciona normal.
