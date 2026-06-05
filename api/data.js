// Persistência na nuvem (Supabase) do estado do dashboard — a "base de dados
// separada". Substitui o antigo modelo em que os lançamentos viviam só no
// localStorage do aparelho (e se perdiam a cada atualização/troca de device).
//
// Tudo fica num único registro (id = 'expenses') de uma tabela `app_state`,
// na coluna `data` (jsonb), guardando o array de lançamentos.
//
// A chave usada aqui é a SERVICE ROLE do Supabase: ela fica SÓ no servidor
// (variáveis de ambiente na Vercel), nunca no navegador. O front-end fala só
// com /api/data, e este arquivo fala com o Supabase.
//
// GET  /api/data            -> { data: [...] }  (ou { data: null } se vazio)
// POST /api/data  { data }  -> grava o array e responde { ok: true }
//
// Se as variáveis de ambiente não estiverem configuradas, responde 503 e o
// front-end cai no fallback local (mesmo comportamento de antes) — assim o app
// nunca quebra, mesmo antes do banco estar pronto.

const ROW_ID = 'expenses';
const TABLE = 'app_state';

function getConfig() {
  let url = process.env.SUPABASE_URL;
  const key =
    process.env.SUPABASE_SERVICE_ROLE_KEY ||
    process.env.SUPABASE_SERVICE_KEY ||
    process.env.SUPABASE_KEY;
  if (!url || !key) return null;
  // Normaliza para apenas o domínio (origin), tolerando valores colados com
  // caminho a mais (ex.: ".../rest/v1") ou barra final — evita montar uma URL
  // inválida como /rest/v1/rest/v1/... (erro PGRST125).
  try { url = new URL(url.trim()).origin; } catch (e) { url = url.trim().replace(/\/+$/, ''); }
  return { url, key };
}

module.exports = async (req, res) => {
  const cfg = getConfig();
  if (!cfg) {
    res.status(503).json({ error: 'no_db' });
    return;
  }

  const base = `${cfg.url}/rest/v1/${TABLE}`;
  const authHeaders = {
    apikey: cfg.key,
    Authorization: `Bearer ${cfg.key}`,
  };

  try {
    if (req.method === 'GET') {
      const r = await fetch(
        `${base}?id=eq.${ROW_ID}&select=data`,
        { headers: authHeaders }
      );
      if (!r.ok) {
        const txt = await r.text().catch(() => '');
        res.status(502).json({ error: 'supabase_read_failed', detail: txt });
        return;
      }
      const rows = await r.json().catch(() => []);
      const data = Array.isArray(rows) && rows[0] ? rows[0].data : null;
      res.status(200).json({ data });
      return;
    }

    if (req.method === 'POST') {
      const body =
        req.body && typeof req.body === 'object'
          ? req.body
          : JSON.parse(req.body || '{}');
      const data = body.data;
      if (!Array.isArray(data)) {
        res.status(400).json({ error: 'invalid_payload' });
        return;
      }

      // Upsert: insere ou atualiza o registro único (id = 'expenses').
      const r = await fetch(`${base}?on_conflict=id`, {
        method: 'POST',
        headers: {
          ...authHeaders,
          'content-type': 'application/json',
          Prefer: 'resolution=merge-duplicates,return=minimal',
        },
        body: JSON.stringify({
          id: ROW_ID,
          data,
          updated_at: new Date().toISOString(),
        }),
      });
      if (!r.ok) {
        const txt = await r.text().catch(() => '');
        res.status(502).json({ error: 'supabase_write_failed', detail: txt });
        return;
      }
      res.status(200).json({ ok: true });
      return;
    }

    res.status(405).json({ error: 'method_not_allowed' });
  } catch (e) {
    res.status(500).json({ error: String((e && e.message) || e) });
  }
};
