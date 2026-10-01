// ═══════════════════════════════════════════════════════════════
// INTEGRAÇÃO KOMMO → TRELLO
// Cada etapa do funil vira uma lista no quadro do Trello e cada lead
// vira um cartão, que acompanha o lead quando ele muda de etapa.
// ═══════════════════════════════════════════════════════════════

const axios = require('axios');
const express = require('express');

const TRELLO_BASE = 'https://api.trello.com/1';
const TRELLO_INTERVALO_MS = 120; // ~8 req/s (limite do Trello: 100 req/10s por token)
const TRELLO_MAX_RETRIES = 4;
const STATUS_GANHO = 142;
const STATUS_PERDIDO = 143;
const MARCADOR_REGEX = /kommo_lead_id:\s*(\d+)/i;

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
const normalizar = s => String(s || '').trim().toLowerCase();

// ═══════════════════════════════════════════════════════════════
// CLIENTE HTTP DO TRELLO (fila + retry em 429)
// ═══════════════════════════════════════════════════════════════

class TrelloClient {
    constructor(apiKey, token) {
        this.authHeader = `OAuth oauth_consumer_key="${apiKey}", oauth_token="${token}"`;
        this.fila = Promise.resolve();
        this.ultimaReq = 0;
    }

    async request(method, path, params, tentativa = 0) {
        const executar = async () => {
            const espera = this.ultimaReq + TRELLO_INTERVALO_MS - Date.now();
            if (espera > 0) await sleep(espera);
            this.ultimaReq = Date.now();
            return axios({ method, url: `${TRELLO_BASE}${path}`, params, headers: { Authorization: this.authHeader } });
        };
        const p = this.fila.then(executar);
        this.fila = p.catch(() => {});
        try {
            return (await p).data;
        } catch (error) {
            if (error.response?.status === 429 && tentativa < TRELLO_MAX_RETRIES) {
                const delay = 1000 * Math.pow(2, tentativa);
                console.warn(`   ⚠️  Trello 429. Retry #${tentativa + 1} em ${delay}ms...`);
                await sleep(delay);
                return this.request(method, path, params, tentativa + 1);
            }
            throw error;
        }
    }
}

// ═══════════════════════════════════════════════════════════════
// CONFIGURAÇÃO (.env)
// ═══════════════════════════════════════════════════════════════

function carregarSyncs(clientes) {
    const syncs = [];
    let n = 1;
    while (process.env[`TRELLO_SYNC${n}_BOARD_ID`]) {
        const subdominio = process.env[`TRELLO_SYNC${n}_SUBDOMINIO`];
        const pipelineId = parseInt(process.env[`TRELLO_SYNC${n}_PIPELINE_ID`]);
        const cliente = clientes.find(c => c.subdominio === subdominio);
        if (!cliente) {
            console.warn(`⚠ TRELLO_SYNC${n}: nenhum CLIENTE com subdomínio "${subdominio}" — sync ignorado.`);
        } else if (!pipelineId) {
            console.warn(`⚠ TRELLO_SYNC${n}: TRELLO_SYNC${n}_PIPELINE_ID ausente ou inválido — sync ignorado.`);
        } else {
            syncs.push({
                id: n,
                cliente,
                pipelineId,
                boardId: process.env[`TRELLO_SYNC${n}_BOARD_ID`],
                incluirFechados: (process.env[`TRELLO_SYNC${n}_INCLUIR_FECHADOS`] || 'sim').toLowerCase() === 'sim',
                webhookSecret: process.env[`TRELLO_SYNC${n}_WEBHOOK_SECRET`] || null,
                fila: Promise.resolve(),
                ultimaExecucao: null
            });
        }
        n++;
    }
    return syncs;
}

// ═══════════════════════════════════════════════════════════════
// LÓGICA DE SINCRONIZAÇÃO
// ═══════════════════════════════════════════════════════════════

function criarSincronizador({ trello, kommoGet }) {
    const usuariosCache = {};

    async function buscarUsuarios(cliente) {
        const c = usuariosCache[cliente.subdominio];
        if (c && Date.now() - c.timestamp < 30 * 60 * 1000) return c.mapa;
        const mapa = {};
        try {
            const r = await kommoGet(cliente.subdominio, cliente.token, '/api/v4/users?limit=250');
            (r.data?._embedded?.users || []).forEach(u => { mapa[u.id] = u.name; });
        } catch (err) { }
        usuariosCache[cliente.subdominio] = { mapa, timestamp: Date.now() };
        return mapa;
    }

    async function buscarLeadsDoFunil(sync) {
        const { cliente, pipelineId } = sync;
        let leads = [], page = 1, hasMore = true;
        while (hasMore) {
            try {
                const r = await kommoGet(cliente.subdominio, cliente.token,
                    `/api/v4/leads?limit=250&page=${page}&filter[pipeline_id][]=${pipelineId}`);
                const lote = r.data?._embedded?.leads || [];
                leads = leads.concat(lote);
                if (lote.length < 250) hasMore = false; else page++;
            } catch (err) {
                if (err.response?.status === 204) hasMore = false; else throw err;
            }
        }
        return leads.filter(l => l.pipeline_id === pipelineId);
    }

    async function buscarLead(sync, leadId) {
        try {
            const r = await kommoGet(sync.cliente.subdominio, sync.cliente.token, `/api/v4/leads/${leadId}`);
            return r.data || null;
        } catch (err) {
            if ([204, 404].includes(err.response?.status)) return null;
            throw err;
        }
    }

    // Garante uma lista no quadro para cada etapa do funil (casando pelo nome).
    async function garantirListas(sync) {
        const r = await kommoGet(sync.cliente.subdominio, sync.cliente.token, `/api/v4/leads/pipelines/${sync.pipelineId}`);
        const statuses = [...(r.data?._embedded?.statuses || [])].sort((a, b) => a.sort - b.sort);

        const listas = await trello.request('get', `/boards/${sync.boardId}/lists`, { filter: 'open', fields: 'name' });
        const porNome = new Map(listas.map(l => [normalizar(l.name), l.id]));

        const mapa = {};
        for (const status of statuses) {
            const fechado = status.id === STATUS_GANHO || status.id === STATUS_PERDIDO;
            if (fechado && !sync.incluirFechados) continue;
            let idLista = porNome.get(normalizar(status.name));
            if (!idLista) {
                const nova = await trello.request('post', '/lists', { name: status.name, idBoard: sync.boardId, pos: 'bottom' });
                idLista = nova.id;
                porNome.set(normalizar(status.name), idLista);
                console.log(`   ➕ Lista "${status.name}" criada no Trello`);
            }
            mapa[status.id] = idLista;
        }
        return mapa;
    }

    // Mapeia lead_id → cartão (inclui arquivados, para não duplicar).
    async function carregarCards(sync) {
        const cards = await trello.request('get', `/boards/${sync.boardId}/cards/all`, { fields: 'name,desc,idList,closed' });
        const porLead = new Map();
        cards.forEach(card => {
            const m = card.desc && card.desc.match(MARCADOR_REGEX);
            if (m) porLead.set(Number(m[1]), card);
        });
        return porLead;
    }

    function montarCard(sync, lead, usuarios) {
        const linhas = [];
        const valor = Number(lead.price) || 0;
        linhas.push(`**Valor:** ${valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}`);
        if (lead.responsible_user_id) linhas.push(`**Responsável:** ${usuarios[lead.responsible_user_id] || `Usuário ${lead.responsible_user_id}`}`);
        if (lead.created_at) linhas.push(`**Criado em:** ${new Date(lead.created_at * 1000).toLocaleDateString('pt-BR')}`);
        const tags = (lead._embedded?.tags || []).map(t => t.name).filter(Boolean);
        if (tags.length) linhas.push(`**Tags:** ${tags.join(', ')}`);
        linhas.push('', `[Abrir no Kommo](https://${sync.cliente.subdominio}.kommo.com/leads/detail/${lead.id})`);
        linhas.push('', '---', `kommo_lead_id: ${lead.id}`);
        return { name: lead.name || `Lead #${lead.id}`, desc: linhas.join('\n') };
    }

    async function upsertCard(sync, lead, listas, cards, usuarios, resumo) {
        const idList = listas[lead.status_id];
        const existente = cards.get(lead.id);
        if (!idList) {
            // Lead fora das etapas sincronizadas (ex.: fechado com INCLUIR_FECHADOS=nao)
            if (existente && !existente.closed) {
                await trello.request('put', `/cards/${existente.id}`, { closed: true });
                existente.closed = true;
                resumo.arquivados++;
            }
            return;
        }
        const { name, desc } = montarCard(sync, lead, usuarios);
        if (!existente) {
            const novo = await trello.request('post', '/cards', { idList, name, desc, pos: 'top' });
            cards.set(lead.id, novo);
            resumo.criados++;
            return;
        }
        const mudancas = {};
        if (existente.idList !== idList) mudancas.idList = idList;
        if (existente.name !== name) mudancas.name = name;
        if (existente.desc !== desc) mudancas.desc = desc;
        if (Object.keys(mudancas).length === 0) { resumo.inalterados++; return; }
        if (mudancas.idList) mudancas.pos = 'top';
        await trello.request('put', `/cards/${existente.id}`, mudancas);
        Object.assign(existente, mudancas);
        resumo.atualizados++;
    }

    async function arquivarCard(card, resumo) {
        if (!card || card.closed) return;
        await trello.request('put', `/cards/${card.id}`, { closed: true });
        card.closed = true;
        resumo.arquivados++;
    }

    // Execuções do mesmo sync são serializadas para evitar cartões duplicados
    // (ex.: webhook chegando durante uma sincronização completa).
    function enfileirar(sync, tipo, fn) {
        const p = sync.fila.then(async () => {
            const inicio = Date.now();
            const resumo = { criados: 0, atualizados: 0, arquivados: 0, inalterados: 0 };
            try {
                await fn(resumo);
                sync.ultimaExecucao = { tipo, ok: true, em: new Date().toISOString(), duracaoMs: Date.now() - inicio, ...resumo };
                console.log(`🔁 Trello sync #${sync.id} (${tipo}): +${resumo.criados} criados, ${resumo.atualizados} atualizados, ${resumo.arquivados} arquivados`);
                return sync.ultimaExecucao;
            } catch (error) {
                const detalhe = error.response ? `${error.response.status} ${JSON.stringify(error.response.data)}` : error.message;
                sync.ultimaExecucao = { tipo, ok: false, em: new Date().toISOString(), erro: detalhe };
                console.error(`❌ Trello sync #${sync.id} (${tipo}): ${detalhe}`);
                throw error;
            }
        });
        sync.fila = p.catch(() => {});
        return p;
    }

    function sincronizarCompleto(sync) {
        return enfileirar(sync, 'completo', async resumo => {
            const listas = await garantirListas(sync);
            const [cards, leads, usuarios] = await Promise.all([
                carregarCards(sync), buscarLeadsDoFunil(sync), buscarUsuarios(sync.cliente)
            ]);
            const idsAtuais = new Set();
            for (const lead of leads) {
                idsAtuais.add(lead.id);
                await upsertCard(sync, lead, listas, cards, usuarios, resumo);
            }
            // Leads excluídos ou movidos para outro funil → arquiva o cartão
            for (const [leadId, card] of cards) {
                if (!idsAtuais.has(leadId)) await arquivarCard(card, resumo);
            }
        });
    }

    function sincronizarLeads(sync, leadIds, removidos = []) {
        return enfileirar(sync, 'webhook', async resumo => {
            const listas = await garantirListas(sync);
            const [cards, usuarios] = await Promise.all([carregarCards(sync), buscarUsuarios(sync.cliente)]);
            for (const id of removidos) await arquivarCard(cards.get(id), resumo);
            for (const id of leadIds) {
                const lead = await buscarLead(sync, id);
                if (!lead || lead.pipeline_id !== sync.pipelineId) await arquivarCard(cards.get(id), resumo);
                else await upsertCard(sync, lead, listas, cards, usuarios, resumo);
            }
        });
    }

    return { sincronizarCompleto, sincronizarLeads };
}

// Extrai IDs do payload de webhook do Kommo (form-urlencoded):
// leads[add|update|status|delete][0][id]=123
function extrairIdsWebhook(body) {
    const leads = body?.leads || {};
    const ids = tipo => Object.values(leads[tipo] || {}).map(l => Number(l?.id)).filter(Boolean);
    const removidos = new Set(ids('delete'));
    const alterados = new Set([...ids('add'), ...ids('update'), ...ids('status')].filter(id => !removidos.has(id)));
    return { alterados: [...alterados], removidos: [...removidos] };
}

// ═══════════════════════════════════════════════════════════════
// ROTAS E AGENDAMENTO
// ═══════════════════════════════════════════════════════════════

function iniciarIntegracaoTrello({ app, clientes, kommoGet }) {
    const apiKey = process.env.TRELLO_API_KEY;
    const token = process.env.TRELLO_TOKEN;
    const syncs = carregarSyncs(clientes);

    app.get('/api/trello/status', (req, res) => {
        res.json({
            configurado: Boolean(apiKey && token),
            syncs: syncs.map(s => ({
                id: s.id, cliente: s.cliente.nome, pipelineId: s.pipelineId, boardId: s.boardId,
                incluirFechados: s.incluirFechados, webhookAtivo: Boolean(s.webhookSecret),
                ultimaExecucao: s.ultimaExecucao
            }))
        });
    });

    if (!apiKey || !token) {
        if (syncs.length) console.warn('⚠ Integração Trello: defina TRELLO_API_KEY e TRELLO_TOKEN no .env.');
        return;
    }
    if (!syncs.length) return;

    const trello = new TrelloClient(apiKey, token);
    const { sincronizarCompleto, sincronizarLeads } = criarSincronizador({ trello, kommoGet });
    const encontrarSync = id => syncs.find(s => s.id === parseInt(id));

    app.post('/api/trello/sincronizar/:syncId', async (req, res) => {
        const sync = encontrarSync(req.params.syncId);
        if (!sync) return res.status(404).json({ erro: 'Sync não encontrado' });
        try { res.json(await sincronizarCompleto(sync)); }
        catch (e) { res.status(502).json(sync.ultimaExecucao); }
    });

    // URL para cadastrar no Kommo: /api/webhooks/kommo/<syncId>?secret=<TRELLO_SYNCn_WEBHOOK_SECRET>
    app.post('/api/webhooks/kommo/:syncId', express.urlencoded({ extended: true }), express.json(), (req, res) => {
        const sync = encontrarSync(req.params.syncId);
        if (!sync || !sync.webhookSecret || req.query.secret !== sync.webhookSecret) return res.sendStatus(404);
        // O Kommo exige resposta rápida; o processamento segue em segundo plano.
        res.sendStatus(200);
        const { alterados, removidos } = extrairIdsWebhook(req.body);
        if (alterados.length || removidos.length) sincronizarLeads(sync, alterados, removidos).catch(() => {});
    });

    const intervaloMin = parseInt(process.env.TRELLO_SYNC_INTERVALO_MIN ?? '15');
    console.log(`🔗 Integração Trello: ${syncs.length} sync(s)${intervaloMin > 0 ? `, a cada ${intervaloMin} min` : ''}`);
    syncs.forEach(sync => {
        console.log(`   • ${sync.cliente.nome} — funil ${sync.pipelineId} → quadro ${sync.boardId}`);
        sincronizarCompleto(sync).catch(() => {});
        if (intervaloMin > 0) setInterval(() => sincronizarCompleto(sync).catch(() => {}), intervaloMin * 60 * 1000);
    });
}

module.exports = { iniciarIntegracaoTrello, criarSincronizador, extrairIdsWebhook, TrelloClient };
