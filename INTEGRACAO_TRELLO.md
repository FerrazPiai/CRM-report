# Integração Kommo → Trello

Espelha um funil do Kommo em um quadro do Trello:

- cada **etapa do funil** vira uma **lista** no quadro (listas que já existem com o mesmo nome são reaproveitadas);
- cada **lead** vira um **cartão** com valor, responsável, data de criação, tags e link para o lead no Kommo;
- quando o lead muda de etapa, o cartão muda de lista;
- leads excluídos ou movidos para outro funil têm o cartão **arquivado** (nunca apagado).

A sincronização é de mão única (Kommo → Trello). A descrição do cartão é sobrescrita pela integração,
então anotações da equipe devem ficar em comentários, checklists ou etiquetas do Trello.

## 1. Credenciais do Trello

1. Acesse <https://trello.com/power-ups/admin>, crie uma Power-Up (pode ser privada) e abra **API key**.
2. Copie a **API key** e clique em **Token** para gerar o token de acesso da conta que tem acesso ao quadro.
3. Pegue o ID do quadro na URL: `https://trello.com/b/<ID_DO_QUADRO>/nome-do-quadro`.

## 2. Configuração no `.env`

O funil é lido com o token de um `CLIENTEn` já cadastrado, encontrado pelo subdomínio:

```env
CLIENTE1_NOME="Nome do cliente"
CLIENTE1_SUBDOMINIO="cassianogueira"
CLIENTE1_TOKEN="token_de_longa_duracao_do_kommo"

TRELLO_API_KEY="..."
TRELLO_TOKEN="..."
TRELLO_SYNC_INTERVALO_MIN=15          # sincronização completa periódica (0 desativa)

TRELLO_SYNC1_SUBDOMINIO="cassianogueira"
TRELLO_SYNC1_PIPELINE_ID="14338679"   # .../leads/pipeline/14338679/
TRELLO_SYNC1_BOARD_ID="ID_DO_QUADRO"
TRELLO_SYNC1_INCLUIR_FECHADOS="sim"   # "nao" = sem listas de Ganho/Perdido
TRELLO_SYNC1_WEBHOOK_SECRET="valor-aleatorio-longo"
```

Para espelhar mais funis, use `TRELLO_SYNC2_*`, `TRELLO_SYNC3_*` etc.

## 3. Tempo real (webhook do Kommo)

Sem webhook, o quadro é atualizado a cada `TRELLO_SYNC_INTERVALO_MIN` minutos. Para refletir as
mudanças em segundos:

1. No Kommo: **Configurações → Integrações → Web hooks**.
2. URL: `https://SEU_DOMINIO/api/webhooks/kommo/1?secret=<TRELLO_SYNC1_WEBHOOK_SECRET>`
3. Eventos: **Lead adicionado**, **Lead alterado**, **Etapa do lead alterada**, **Lead excluído**.

O servidor precisa estar acessível publicamente por HTTPS.

## 4. Endpoints

| Método | Rota | Uso |
| --- | --- | --- |
| `GET` | `/api/trello/status` | Configuração dos syncs e resultado da última execução |
| `POST` | `/api/trello/sincronizar/:syncId` | Força uma sincronização completa agora |
| `POST` | `/api/webhooks/kommo/:syncId?secret=...` | Recebe os webhooks do Kommo |
