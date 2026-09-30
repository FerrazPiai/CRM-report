# Clusterização da Base de Leads — Camiseta Express

Pipeline que junta **todas** as planilhas de contatos da pasta do Google Drive
(Locaweb, formulários do site/Elementor, agendas/WhatsApp e exportações do Outlook),
higieniza, remove duplicados, segmenta por **empresa/segmento** e por **temperatura**
e gera uma planilha única, separada por abas, pronta para subir no **RD Station**.

> ⚠️ **LGPD / repositório público** — este repositório é público. Os dados pessoais
> (`dados_brutos/` e `saida/`) estão no `.gitignore` e **nunca** devem ser commitados.
> Aqui fica só o código e a metodologia; a planilha final é compartilhada por fora
> (Drive privado).

## Como rodar

```bash
pip install pandas openpyxl xlsxwriter
# 1. Baixe a pasta do Drive e coloque os arquivos em:
#    dados_brutos/raiz/        -> planilhas soltas da pasta principal
#    dados_brutos/locaweb/     -> "Camiseta Express Leads Lista Contatos Email Mkt Locaweb"
#    dados_brutos/site_forms/  -> "Camiseta Express Leads Site Forms"
# 2. Rode:
python3 processar_base.py
# 3. Resultado em saida/
```

Nomes esperados em `dados_brutos/raiz/` (Google Sheets exportados como `.xlsx`):
`ContatosCamex24-07.csv`, `ContatosDebora21-09.xlsx`, `ContatosSilmara.xlsx`,
`EmailsDebora22-09.xlsx`, `EmailsRose23-09.xlsx`, `Leads_test_chat.xlsx`,
`Lista_de_emails_submissoes.xlsx`, `Atualizacao_Cadastro_B2B.xlsx`.
Os CSVs da Locaweb e do Elementor são lidos automaticamente (qualquer nome).
Opcional: `dados_brutos/emails_internos.txt` com regex (1 por linha) de e-mails da equipe usados
em testes, para excluí-los da base.
Arquivos duplicados (mesmo conteúdo) são detectados por hash e ignorados.

## Saídas (`saida/`)

| Arquivo | Conteúdo |
|---|---|
| `Base_Clusterizada_Camiseta_Express.xlsx` | Planilha única com todas as abas (abaixo) |
| `rdstation_marketing_contatos.csv` | Contatos com e-mail válido, sem descadastrados e sem bounce — importação no RD Station Marketing |
| `rdstation_crm_contatos.csv` | Todos os contatos ativos (inclui só-telefone) — RD Station CRM |
| `rdstation_crm_empresas.csv` | 1 linha por empresa — RD Station CRM (Empresas) |

### Abas da planilha

- **LEIA-ME** — regras de tratamento e como importar.
- **Resumo** — KPIs, Segmento × Temperatura, Período × Temperatura, contatos por origem.
- **Base Completa** — todos os contatos únicos ativos (1 linha por pessoa).
- **Leads Quentes / Mornos / Frios** — base separada por temperatura.
- **Empresas (Contas)** — 1 linha por empresa com nº de contatos e temperatura da conta.
- **Uma aba por segmento**: Governo e Setor Público · Educação e Escolas · Agências Mkt e Publicidade ·
  Brindes e Promocionais · Gráficas e Impressão · Eventos e Buffet · Franquias · Alimentação e Bebidas ·
  Condomínios e Imobiliárias · Saúde e Farmacêutico · Financeiro e Seguros · Esportes e Academias ·
  ONGs, Associações e Igrejas · Tecnologia e Telecom · Indústria e Energia · Varejo, Moda e Atacado ·
  Serviços, Logística e RH · Empresas - Outros Segmentos · Pessoa Física - Sem Empresa.
- **Descadastrados (NÃO importar)** — opt-out da Locaweb.
- **Correções de E-mail** — log original → corrigido.
- **Removidos** — internos, testes, spam, e-mails de sistema e registros sem contato válido.
- **Arquivos Processados** — auditoria de cada arquivo lido.

## Regras de tratamento

**Telefones** — formato `+55 (DDD) número`. Sem DDD → **DDD 11** (operação em SP), sinalizado na
coluna `DDD 11 assumido?`. Celular antigo de 8 dígitos ganha o 9º dígito. Remove código de país,
prefixo de operadora (`0XX15`, `015 11`…) e zero à esquerda. Descarta DDD inexistente, 0800,
números curtos ou repetidos (`99999999`). Internacionais mantidos com `+código`.

**E-mails** — minúsculas, sem espaços/aspas/`mailto:`, vírgula → ponto, `@` faltante ou duplicado,
extensões erradas (`.con`, `.cm`, `.comc`, `.com.b`, `.combr`, `.br.com`, `gmail.com.br`…),
provedores digitados errado (`gmial`, `gamil`, `hotmial`, `hormail`, `outlok`, `yahooo`…) e domínio
corporativo com **1 letra faltando** em relação a um domínio frequente na base (ex.: `erofarma` →
`eurofarma`). Placeholders (`teste@`, `email@`, `naotem@`, `xxx@`) são descartados. Tudo logado na
aba *Correções de E-mail*.

**Duplicados** — união por e-mail **ou** celular (union-find entre todas as fontes). Telefones
compartilhados por 3+ e-mails (PABX) e fixos compartilhados não unificam. Na linha final ficam o
e-mail principal (preferência corporativo), celular, demais e-mails/telefones, todas as origens,
listas, formulários e vendedoras.

**Limpeza** — removidos e-mails `@camisetaexpress.com.br` e da equipe de desenvolvimento, registros de
teste, spam de formulário (cirílico/chinês, domínios `.top/.ru`, mensagens em inglês com nome de robô,
propaganda) — o spam é propagado para todas as fontes do mesmo e-mail —, e-mails automáticos
(`no-reply`, `mailer-daemon`…) e registros sem nenhum contato válido.

**Segmento** (ordem de prioridade) — 1) TLD institucional (`.gov.br`, `.jus.br`, `.edu.br`) →
2) palavras-chave e marcas no domínio/nome da empresa → 3) lista de segmento da Locaweb (agências,
brindes, impressão digital, buffet, franquias, alimentícias, condomínios, escola de inglês,
atacadistas) → 4) formulários de uniforme escolar → 5) demais TLDs (`.org`, `.io`…).
Sem empresa identificável (e-mail pessoal e sem empresa informada) → *Pessoa Física - Sem Empresa*.
Contatos só-telefone cujo nome na agenda tem marcador de empresa (`Ltda`, `Escola`, `Buffet`,
`Gráfica`…) viram empresa. **A classificação é automática — revise "Empresas - Outros Segmentos".**

**Data de entrada** — data mais antiga do contato: envio de formulário (site/B2B) ou importação
na Locaweb. As datas da Locaweb são **lotes de importação** (07, 18 e 22/10/2024 e 30/07/2026), então
indicam "está na base pelo menos desde". Contatos só de agenda/Outlook (sem data por registro) usam a
data da exportação. Coluna `Origem da data de entrada` indica a fonte.

**Temperatura (score 0–100)**

| Critério | Pontos |
|---|---|
| Última interação real (formulário) ≤ 90 dias / ≤ 180 / ≤ 365 / ≤ 2 anos / > 2 anos | 45 / 35 / 25 / 12 / 5 |
| Lead inbound (site ou chat) | +20 |
| Trocou e-mails com vendedora no Outlook (3+ e-mails) | +15 (+10) |
| Salvo na agenda/WhatsApp de vendedora | +12 |
| Clicou em e-mail mkt (ou só visualizou) | +15 (+5) |
| Aparece em várias fontes | +5 por fonte extra (máx. 10) |
| Perfil B2B | +5 |
| E-mail com bounce na Locaweb | −10 |

**Quente ≥ 55 · Morno 30–54 · Frio < 30 · Descadastrado = opt-out.**

## Importação no RD Station

- **Marketing**: `rdstation_marketing_contatos.csv`. Mapear Nome, Email, Celular, Telefone, Empresa,
  Cidade, Estado; a coluna `Tags` já traz `segmento, lead-temperatura, entrada-periodo, origem-*, resp-*`
  (Segmento/Temperatura/Score podem virar campos personalizados). Não importar descadastrados.
- **CRM**: importar primeiro `rdstation_crm_empresas.csv` (Empresas) e depois `rdstation_crm_contatos.csv`
  (Contatos, vinculando pela coluna Empresa). Contatos sem e-mail só entram no CRM.
