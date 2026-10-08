// ═══════════════════════════════════════════════════════════════
// INTEGRAÇÃO KOMMO (GANHO) → TRELLO (PRODUÇÃO)
// ═══════════════════════════════════════════════════════════════
//
// Quando um lead do funil configurado entra na etapa "Ganho" (status 142
// no Kommo), o webhook cria um cartão na lista de Produção do Trello com
// os dados do lead e registra o link do cartão como nota no lead.

const axios = require('axios');
const crypto = require('crypto');
const express = require('express');

const STATUS_GANHO = 142;
const TRELLO_API = 'https://api.trello.com/1';
const MARCADOR = id => `[kommo-lead:${id}]`;

function carregarConfig(clientes) {
    const subdominio = process.env.KOMMO_TRELLO_SUBDOMINIO;
    const clienteMesmoSub = clientes.find(c => c.subdominio === subdominio);
    return {
        subdominio,
        token: process.env.KOMMO_TRELLO_TOKEN || clienteMesmoSub?.token,
        pipelineId: parseInt(process.env.KOMMO_TRELLO_PIPELINE_ID) || null,
        trelloKey: process.env.TRELLO_API_KEY,
        trelloToken: process.env.TRELLO_TOKEN,
        trelloListId: process.env.TRELLO_LISTA_PRODUCAO_ID,
        trelloEtiquetas: process.env.TRELLO_ETIQUETAS_IDS || '',
        segredo: process.env.WEBHOOK_SEGREDO
    };
}

function segredoValido(recebido, esperado) {
    if (!recebido || !esperado) return false;
    const a = Buffer.from(String(recebido));
    const b = Buffer.from(String(esperado));
    return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function formatarMoeda(valor) {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(valor || 0);
}

function formatarCampo(campo) {
    const tiposData = ['date', 'date_time', 'birthday'];
    return (campo.values || []).map(v => {
        if (tiposData.includes(campo.field_type) && typeof v.value === 'number') {
            return new Date(v.value * 1000).toLocaleDateString('pt-BR');
        }
        return v.value;
    }).filter(v => v !== null && v !== undefined && v !== '').join(', ');
}

function valorPorCodigo(entidade, codigo) {
    const campo = (entidade?.custom_fields_values || []).find(c => c.field_code === codigo);
    return campo ? formatarCampo(campo) : '';
}

// Kommo envia o webhook como form-urlencoded: leads[status][0][id]=...
// Dependendo da quantidade, o parser devolve array ou objeto indexado.
function extrairLeadsGanhos(body, config) {
    const grupos = [body?.leads?.status, body?.leads?.add];
    const leads = grupos.flatMap(g => (g ? Object.values(g) : []));
    return leads
        .filter(l => parseInt(l.status_id) === STATUS_GANHO)
        .filter(l => !config.pipelineId || parseInt(l.pipeline_id) === config.pipelineId)
        .map(l => parseInt(l.id))
        .filter(Boolean);
}

function registrarIntegracaoTrello(app, { clientes, rateLimiter }) {
    const config = carregarConfig(clientes);
    const eventos = [];
    const emAndamento = new Set();

    const faltando = ['subdominio', 'token', 'pipelineId', 'trelloKey', 'trelloToken', 'trelloListId', 'segredo']
        .filter(k => !config[k]);
    if (faltando.length) {
        console.log(`🟡 Integração Trello inativa — faltando: ${faltando.join(', ')}`);
    } else {
        console.log(`🟢 Integração Kommo→Trello ativa (${config.subdominio}, funil ${config.pipelineId})`);
    }

    function registrarEvento(evento) {
        eventos.unshift({ ...evento, em: new Date().toISOString() });
        eventos.length = Math.min(eventos.length, 50);
    }

    function kommo(method, urlPath, data) {
        return rateLimiter.request(config.subdominio, {
            method,
            url: `https://${config.subdominio}.kommo.com${urlPath}`,
            headers: { 'Authorization': `Bearer ${config.token}` },
            data
        });
    }

    function trello(method, urlPath, params = {}) {
        return axios({
            method,
            url: `${TRELLO_API}${urlPath}`,
            params: { ...params, key: config.trelloKey, token: config.trelloToken }
        });
    }

    async function buscarDetalhesLead(leadId) {
        const { data: lead } = await kommo('get', `/api/v4/leads/${leadId}?with=contacts`);

        const contatos = lead._embedded?.contacts || [];
        const principal = contatos.find(c => c.is_main) || contatos[0];
        const contato = principal ? (await kommo('get', `/api/v4/contacts/${principal.id}`)).data : null;

        const empresaRef = lead._embedded?.companies?.[0];
        const empresa = empresaRef
            ? await kommo('get', `/api/v4/companies/${empresaRef.id}`).then(r => r.data).catch(() => null)
            : null;

        const responsavel = lead.responsible_user_id
            ? await kommo('get', `/api/v4/users/${lead.responsible_user_id}`).then(r => r.data).catch(() => null)
            : null;

        return { lead, contato, empresa, responsavel };
    }

    function montarCartao({ lead, contato, empresa, responsavel }) {
        const nomeLeadPadrao = !lead.name || /^(lead|сделка) #\d+$/i.test(lead.name);
        const titulo = (nomeLeadPadrao && contato?.name) ? contato.name : (lead.name || `Lead #${lead.id}`);

        const linhas = [
            `**Valor:** ${formatarMoeda(lead.price)}`,
            contato && `**Contato:** ${contato.name}`,
            valorPorCodigo(contato, 'PHONE') && `**Telefone:** ${valorPorCodigo(contato, 'PHONE')}`,
            valorPorCodigo(contato, 'EMAIL') && `**E-mail:** ${valorPorCodigo(contato, 'EMAIL')}`,
            empresa && `**Empresa:** ${empresa.name}`,
            responsavel && `**Vendedor:** ${responsavel.name}`,
            lead.closed_at && `**Fechado em:** ${new Date(lead.closed_at * 1000).toLocaleDateString('pt-BR')}`
        ].filter(Boolean);

        const campos = (lead.custom_fields_values || [])
            .map(c => ({ nome: c.field_name, valor: formatarCampo(c) }))
            .filter(c => c.valor);
        if (campos.length) {
            linhas.push('', '**Campos do lead**');
            campos.forEach(c => linhas.push(`- ${c.nome}: ${c.valor}`));
        }

        linhas.push('', `[Abrir no Kommo](https://${config.subdominio}.kommo.com/leads/detail/${lead.id})`, '', MARCADOR(lead.id));

        // Trello aceita até 16384 caracteres na descrição
        let desc = linhas.join('\n');
        if (desc.length > 16000) desc = `${desc.slice(0, 15900)}\n…\n\n${MARCADOR(lead.id)}`;
        return { name: titulo.slice(0, 500), desc };
    }

    async function buscarCartaoExistente(leadId) {
        const { data: lista } = await trello('get', `/lists/${config.trelloListId}`, { fields: 'idBoard' });
        const { data: cartoes } = await trello('get', `/boards/${lista.idBoard}/cards/all`, { fields: 'desc,shortUrl' });
        return cartoes.find(c => (c.desc || '').includes(MARCADOR(leadId)));
    }

    async function sincronizarLead(leadId) {
        if (emAndamento.has(leadId)) return { leadId, resultado: 'ignorado', motivo: 'já em processamento' };
        emAndamento.add(leadId);
        try {
            const existente = await buscarCartaoExistente(leadId);
            if (existente) {
                const r = { leadId, resultado: 'ignorado', motivo: 'cartão já existe', cartao: existente.shortUrl };
                registrarEvento(r);
                return r;
            }

            const detalhes = await buscarDetalhesLead(leadId);
            const cartao = montarCartao(detalhes);
            const params = { idList: config.trelloListId, pos: 'top', ...cartao };
            if (config.trelloEtiquetas) params.idLabels = config.trelloEtiquetas;
            const { data: criado } = await trello('post', '/cards', params);
            console.log(`🟩 Trello: cartão criado para lead ${leadId} → ${criado.shortUrl}`);

            await kommo('post', '/api/v4/leads/notes', [{
                entity_id: leadId,
                note_type: 'common',
                params: { text: `Enviado para Produção no Trello: ${criado.shortUrl}` }
            }]).catch(e => console.warn(`   ⚠️  Não foi possível anotar o lead ${leadId}: ${e.response?.status || e.message}`));

            const r = { leadId, resultado: 'criado', cartao: criado.shortUrl };
            registrarEvento(r);
            return r;
        } catch (e) {
            const detalhe = e.response ? `${e.response.status} ${JSON.stringify(e.response.data).slice(0, 300)}` : e.message;
            console.error(`❌ Trello: falha no lead ${leadId}: ${detalhe}`);
            const r = { leadId, resultado: 'erro', motivo: detalhe };
            registrarEvento(r);
            return r;
        } finally {
            emAndamento.delete(leadId);
        }
    }

    function exigirConfig(req, res, next) {
        if (faltando.length) return res.status(503).json({ erro: 'Integração não configurada', faltando });
        next();
    }

    function exigirSegredo(req, res, next) {
        const recebido = req.params.segredo || req.query.segredo || req.get('x-segredo');
        if (!segredoValido(recebido, config.segredo)) return res.status(401).json({ erro: 'Segredo inválido' });
        next();
    }

    // Webhook configurado no Kommo. Responde na hora (o Kommo desativa
    // webhooks lentos) e processa os leads em segundo plano.
    app.post('/webhook/kommo-ganho/:segredo',
        exigirConfig, exigirSegredo,
        express.urlencoded({ extended: true, limit: '2mb' }),
        express.json({ limit: '2mb' }),
        (req, res) => {
            res.sendStatus(200);
            const subRecebido = req.body?.account?.subdomain;
            if (subRecebido && subRecebido !== config.subdominio) return;
            const leads = extrairLeadsGanhos(req.body, config);
            if (leads.length) console.log(`📥 Webhook Kommo: ${leads.length} lead(s) ganho(s): ${leads.join(', ')}`);
            leads.reduce((p, id) => p.then(() => sincronizarLead(id)), Promise.resolve());
        });

    // Envio manual de um lead (teste ou reprocessamento)
    app.post('/api/integracao/trello/sincronizar/:leadId', exigirConfig, exigirSegredo, async (req, res) => {
        const leadId = parseInt(req.params.leadId);
        if (!leadId) return res.status(400).json({ erro: 'leadId inválido' });
        res.json(await sincronizarLead(leadId));
    });

    // Lista quadros e listas do Trello para descobrir o ID da lista "Produção"
    app.get('/api/integracao/trello/listas', exigirSegredo, async (req, res) => {
        if (!config.trelloKey || !config.trelloToken) {
            return res.status(503).json({ erro: 'Defina TRELLO_API_KEY e TRELLO_TOKEN' });
        }
        try {
            const { data } = await trello('get', '/members/me/boards', {
                fields: 'name,url', filter: 'open', lists: 'open', list_fields: 'name'
            });
            res.json(data.map(b => ({ quadro: b.name, url: b.url, listas: (b.lists || []).map(l => ({ id: l.id, nome: l.name })) })));
        } catch (e) {
            res.status(502).json({ erro: e.response?.data || e.message });
        }
    });

    app.get('/api/integracao/trello/status', exigirSegredo, (req, res) => {
        res.json({ ativa: !faltando.length, faltando, subdominio: config.subdominio, pipelineId: config.pipelineId, eventos });
    });
}

module.exports = { registrarIntegracaoTrello, extrairLeadsGanhos };
