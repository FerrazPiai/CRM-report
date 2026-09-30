# Clusterização da base de leads — Camiseta Express

Script que junta todas as planilhas de leads da pasta do cliente no Google Drive, limpa, remove duplicados entre as fontes e separa a base em páginas (abas) com tags prontas para filtros e disparos.

## Como rodar

1. Baixe a pasta do cliente no Drive (botão *Fazer download*) e descompacte, por exemplo em `dados/`.
2. Instale as dependências e rode:

```bash
pip install -r scripts/leads/requirements.txt
python3 scripts/leads/clusterizar_leads.py --entrada dados/ --saida saida/ --data-referencia 2026-09-30 \
    [--equipe equipe_interna.csv]
```

`--equipe` (opcional) recebe uma planilha com os e-mails e telefones da equipe da Camiseta Express. Todos eles vão para a Supressão. O modelo está em `equipe_interna_modelo.csv`.

Saída:
- `saida/Camiseta_Express_Leads_Clusterizados.xlsx`: planilha com as abas descritas abaixo
- `saida/disparos/*.csv` e `saida/disparos.zip`: listas por página, uma de e-mail e uma de WhatsApp (sem supressão)

> Os arquivos de dados (planilhas de entrada e de saída) têm dados pessoais e **não** devem ser commitados. O `.gitignore` da pasta já os bloqueia.

## Fontes reconhecidas automaticamente

| Fonte | Como é identificada | O que traz |
|---|---|---|
| Formulários do site (Elementor) | coluna `Nome do Formulário (ID)` | nome, e-mail, telefone, empresa, interesse, mensagem, **data do envio** |
| Listas Locaweb (e-mail mkt) | colunas `email`, `status`, `data de cadastro` | status (ativo/bounce/opt-out), aberturas, cliques, vertical (nome da lista) |
| Histórico Outlook (`Emails<Vendedora>…`) | arquivo começando com `Emails` | nº de e-mails trocados por endereço (indicador de relacionamento) |
| Agendas WhatsApp/Outlook (`Contatos<Vendedora>…`) | arquivo começando com `Contatos` | nome + telefones, carteira da vendedora. `ContatosCamex` é o **WhatsApp da Rose**; o número é vinculado à conta da loja, então tem clientes da loja misturados |
| Demais planilhas (chat do site, Atualização B2B, lista de submissões) | colunas de e-mail/telefone | contatos adicionais |

## Regras de limpeza

- **Telefone:** só dígitos. Remove +55, o 0 e o código de operadora. **Número sem DDD recebe 11** (operação em SP). Celular antigo com 8 dígitos ganha o 9. Sai formatado como `(11) 91234-5678` e, para WhatsApp, como `5511912345678`.
- **E-mail:** minúsculo, validado, com correção de erros comuns (`gmail.con`, `hotmial.com`, `hormail.com`…).
- **Deduplicação:** registros de todas as fontes viram um único contato quando têm o mesmo e-mail **ou** o mesmo telefone. Um telefone usado por 4 ou mais e-mails diferentes não é usado para juntar registros.
- **Empresa:** primeiro pelo domínio do e-mail corporativo, depois pelo campo *Empresa* do formulário e, por último, pela marca citada no nome do contato na agenda (ex.: "Paulo Carrefour").
- **Removidos da prospecção (aba Supressão):** opt-out na Locaweb, e-mail com bounce e sem telefone, spam/robôs de formulário, testes e endereços automáticos (Odoo, Mailchimp, DocuSign, no-reply). Também saem as pessoas da equipe: domínio `*camisetaexpress*`, nome com "Camiseta Express"/"Camex" e a lista passada em `--equipe`.

## Clusterização

**Núcleo:** são os dois núcleos de venda da cliente. **Acima de 100 peças o pedido vai para a Fábrica; até 100 peças, para a Loja**, as mesmas faixas do formulário do site.

1. Se o lead informou quantidade (formulário ou mensagem), só ela decide, mesmo que seja uma empresa.
2. Sem quantidade: empresa identificada, lista setorial B2B ou revenda de uniforme vão para **Fábrica**. Pessoa física ou lista da loja vão para **Loja**. O resto fica em **A qualificar**.

A coluna `Critério do núcleo` mostra qual regra foi usada em cada contato.

**Setor:** usado nas réguas por segmento (ex.: supermercados, agências). A ordem é: lista de empresas classificadas à mão (`_SETORES_MANUAIS`), depois a lista setorial da Locaweb, depois palavras no nome ou domínio da empresa (`SETOR_POR_PALAVRA`).

**Temperatura:** usa a data de entrada e os indicadores disponíveis.

| Temperatura | Regra |
|---|---|
| 1. Quente | pediu orçamento/WhatsApp pelo site nos últimos 90 dias (proposta em aberto) |
| 2. Relacionamento ativo | 5+ e-mails trocados com a equipe no Outlook |
| 3. Morno | solicitou entre 91 e 365 dias, ou 1 a 4 e-mails trocados, ou clicou em e-mail mkt, ou entrou na base há menos de 1 ano |
| 4. Frio / Reativação | nenhum sinal nos últimos 12 meses |
| 5. Não contatar | supressão |

**Páginas (abas):** cada contato fica em uma só página, para não receber disparo em duplicidade.

| Aba | Conteúdo |
|---|---|
| Contas-Chave | empresas com 5+ contatos, 20+ e-mails trocados ou marcas estratégicas (Carrefour, Safra, Vigor, Atacadão, Nestlé…), ordenadas por empresa |
| Fábrica - Corporativo | núcleo Fábrica: demais empresas finais e listas verticais (alimentícias, atacadistas, franquias, condomínios, buffets, escolas de idiomas) |
| Agências-Brindes-Gráficas | canal indireto: agências (inclui ABAP), promocionais/brindes, gráficas |
| Uniforme Escolar | formulários de uniforme escolar (revenda, varejo, kit municipal) |
| Loja - Varejo | núcleo Loja: pedidos de até 100 peças (inclusive de empresas) e pessoa física |
| A Qualificar | sem sinal de empresa ou intenção, sobretudo agendas de WhatsApp |
| Supressão | não disparar |

Abas **Conta - <empresa>**: as 20 empresas com mais contatos aptos ganham uma aba própria, para disparo específico. É só uma visão: o mesmo contato continua na sua página.

Abas extras: **Resumo** (tabelas por página, temperatura, núcleo, coorte de entrada e fonte), **Base Consolidada**, **Empresas** (uma linha por empresa) e **Dicionário**.

**Score (0–100):** soma recência da solicitação, nº de solicitações, e-mails trocados, engajamento Locaweb, e-mail corporativo, porte da conta, volume estimado e WhatsApp disponível. Serve para ordenar os contatos dentro de cada aba.

**Tags:** ficam na coluna `Tags`, separadas por `; ` no formato `grupo:valor`: `nucleo:`, `temp:`, `pagina:`, `empresa:`, `setor:`, `volume:`, `vertical:`, `conta:`, `entrada:AAAA-Tn`, `fonte:`, `carteira:`, `interesse:`. Há também os marcadores `pediu_orcamento`, `historico_outlook`, `email_corporativo`, `canal:email`, `canal:whatsapp`, `ddd_assumido_11`, `optout` e `email_bounce`.

## Ajustes comuns

No topo do script:
- `LIMITE_FABRICA`: corte de peças entre Fábrica e Loja (padrão 100)
- `QTD_ABAS_CONTAS`: quantas empresas ganham aba própria (padrão 20)
- `EMAILS_INTERNOS` / `EMAILS_INTERNOS_A_CONFIRMAR`: endereços da equipe/parceiros que não são leads
- `_SETORES_MANUAIS` / `SETOR_POR_PALAVRA`: classificação de setor
- `ALIAS_EMPRESAS`: nome de exibição por raiz de domínio (ex.: `kcc` → Kimberly-Clark)
- `MARCAS_EM_NOME`: marcas procuradas no nome dos contatos das agendas
- `LISTAS_LOCAWEB`: vertical, carteira e núcleo de cada lista da Locaweb
- `DATAS_SNAPSHOT`: data de entrada das planilhas sem data por linha
