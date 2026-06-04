// Proxy serverless (Vercel) para gerar insights com a IA da Claude.
// A chave fica no servidor (variável de ambiente ANTHROPIC_API_KEY), nunca no
// navegador. Se a chave não estiver configurada, responde 503 e o front-end
// usa a análise local como fallback.
module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'method_not_allowed' });
    return;
  }
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) {
    res.status(503).json({ error: 'no_key' });
    return;
  }
  try {
    const body = req.body && typeof req.body === 'object' ? req.body : JSON.parse(req.body || '{}');
    const prompt = body.prompt;
    if (!prompt) {
      res.status(400).json({ error: 'no_prompt' });
      return;
    }
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify({
        model: process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-6',
        max_tokens: 1500,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    const data = await r.json();
    if (!r.ok) {
      res.status(r.status).json({ error: (data && data.error && data.error.message) || 'api_error' });
      return;
    }
    const text = (data.content || []).map(b => b.text || '').join('');
    res.status(200).json({ text });
  } catch (e) {
    res.status(500).json({ error: String((e && e.message) || e) });
  }
};
