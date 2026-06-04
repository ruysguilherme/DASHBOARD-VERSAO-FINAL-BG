// Proxy serverless (Vercel) para gerar insights com a IA do Google Gemini
// (camada gratuita). A chave fica no servidor (variável de ambiente
// GEMINI_API_KEY), nunca no navegador. Se a chave não estiver configurada,
// responde 503 e o front-end usa a análise local como fallback.
//
// Tenta uma lista de modelos em ordem: se um retornar cota esgotada (429) ou
// não encontrado (404), passa para o próximo — alguns projetos têm free tier
// em um modelo e não em outro.
module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'method_not_allowed' });
    return;
  }
  const key = process.env.GEMINI_API_KEY;
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

    // Ordem de tentativa: o modelo configurado (se houver) primeiro, depois
    // uma cadeia de fallback.
    const fallback = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-2.5-flash-lite', 'gemini-1.5-flash'];
    const models = [];
    if (process.env.GEMINI_MODEL) models.push(process.env.GEMINI_MODEL);
    fallback.forEach(m => { if (!models.includes(m)) models.push(m); });

    let lastError = 'api_error';
    let lastStatus = 502;

    for (const model of models) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
      const r = await fetch(url, {
        method: 'POST',
        headers: { 'content-type': 'application/json', 'x-goog-api-key': key },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { maxOutputTokens: 1500, temperature: 0.7 },
        }),
      });
      const data = await r.json().catch(() => null);

      if (r.ok) {
        const parts = (((data || {}).candidates || [])[0] || {}).content;
        const text = parts && parts.parts ? parts.parts.map(p => p.text || '').join('') : '';
        if (text.trim()) {
          res.status(200).json({ text, model });
          return;
        }
        lastError = 'empty_response';
        lastStatus = 502;
        continue;
      }

      lastError = (data && data.error && data.error.message) || 'api_error';
      lastStatus = r.status;
      // Chave inválida / requisição malformada: não adianta tentar outros modelos.
      if (r.status === 400 || r.status === 401 || r.status === 403) break;
      // 429 (cota) e 404 (modelo indisponível): tenta o próximo modelo.
    }

    res.status(lastStatus).json({ error: lastError });
  } catch (e) {
    res.status(500).json({ error: String((e && e.message) || e) });
  }
};
