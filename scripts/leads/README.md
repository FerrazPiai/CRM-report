# Clusterização da base de leads — Camiseta Express

Script que junta todas as planilhas de leads da pasta do cliente no Google Drive, limpa, remove duplicados entre as fontes e separa a base em páginas (abas) com tags prontas para filtros e disparos.

## Como rodar

1. Baixe a pasta do cliente no Drive (botão *Fazer download*) e descompacte, por exemplo em `dados/`.
2. Instale as dependências e rode:

```bash
pip install -r scripts/leads/requirements.txt
python3 scripts/leads/clusterizar_leads.py --entrada dados/ --saida saida/ --data-referencia 2026-09-30
```

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
| Agendas WhatsApp/Outlook (`Contatos<Vendedora>…`) | arquivo começando com `Contatos` | nome + telefones, carteira da vendedora |
| Demais planilhas (chat do site, Atualização B2B, lista de submissões) | colunas de e-mail/telefone | contatos adicionais |

## Regras de limpeza

- **Telefone:** só dígitos. Remove +55, o 0 e o código de operadora. **Número sem DDD recebe 11** (operação em SP). Celular antigo com 8 dígitos ganha o 9. Sai formatado como `(11) 91234-5678` e, para WhatsApp, como `5511912345678`.
- **E-mail:** minúsculo, validado, com correção de erros comuns (`gmail.con`, `hotmial.com`, `hormail.com`…).
- **Deduplicação:** registros de todas as fontes viram um único contato quando têm o mesmo e-mail **ou** o mesmo telefone. Um telefone usado por 4 ou mais e-mails diferentes não é usado para juntar registros.
- **Empresa:** primeiro pelo domínio do e-mail corporativo, depois pelo campo *Empresa* do formulário e, por último, pela marca citada no nome do contato na agenda (ex.: "Paulo Carrefour").
- **Removidos da prospecção (aba Supressão):** opt-out na Locaweb, e-mail com bounce e sem telefone, spam/robôs de formulário, endereços internos/sistema (`*camisetaexpress*`, Odoo, no-reply) e testes.

## Clusterização

**Núcleo:** *B2B* (fabricação para empresas), *Loja Física* (pessoa física e projetos menores) ou *A qualificar*.

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
| B2B Corporativo | demais empresas finais e listas verticais (alimentícias, atacadistas, franquias, condomínios, buffets, escolas de idiomas) |
| Agências-Brindes-Gráficas | canal indireto: agências (inclui ABAP), promocionais/brindes, gráficas |
| Uniforme Escolar | formulários de uniforme escolar (revenda, varejo, kit municipal) |
| Loja Física - Varejo | pessoa física / projetos menores |
| A Qualificar | sem sinal de empresa ou intenção, sobretudo agendas de WhatsApp |
| Supressão | não disparar |

Abas extras: **Resumo** (tabelas por página, temperatura, núcleo, coorte de entrada e fonte), **Base Consolidada**, **Empresas** (uma linha por empresa) e **Dicionário**.

**Score (0–100):** soma recência da solicitação, nº de solicitações, e-mails trocados, engajamento Locaweb, e-mail corporativo, porte da conta, volume estimado e WhatsApp disponível. Serve para ordenar os contatos dentro de cada aba.

**Tags:** ficam na coluna `Tags`, separadas por `; ` no formato `grupo:valor`: `nucleo:`, `temp:`, `pagina:`, `vertical:`, `conta:`, `entrada:AAAA-Tn`, `fonte:`, `carteira:`, `interesse:`. Há também os marcadores `pediu_orcamento`, `historico_outlook`, `email_corporativo`, `canal:email`, `canal:whatsapp`, `ddd_assumido_11`, `optout` e `email_bounce`.

## Ajustes comuns

No topo do script:
- `EMAILS_INTERNOS`: endereços da equipe/parceiros que não são leads
- `ALIAS_EMPRESAS`: nome de exibição por raiz de domínio (ex.: `kcc` → Kimberly-Clark)
- `MARCAS_EM_NOME`: marcas procuradas no nome dos contatos das agendas
- `LISTAS_LOCAWEB`: vertical, carteira e núcleo de cada lista da Locaweb
- `DATAS_SNAPSHOT`: data de entrada das planilhas sem data por linha
