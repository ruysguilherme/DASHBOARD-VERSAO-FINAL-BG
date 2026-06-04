// Proxy serverless (Vercel) para gerar insights com a IA do Google Gemini
// (camada gratuita). A chave fica no servidor (variável de ambiente
// GEMINI_API_KEY), nunca no navegador. Se a chave não estiver configurada,
// responde 503 e o front-end usa a análise local como fallback.
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
    const model = process.env.GEMINI_MODEL || 'gemini-2.0-flash';
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
    const r = await fetch(url, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-goog-api-key': key,
      },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 1500, temperature: 0.7 },
      }),
    });
    const data = await r.json();
    if (!r.ok) {
      res.status(r.status).json({ error: (data && data.error && data.error.message) || 'api_error' });
      return;
    }
    const text = (((data.candidates || [])[0] || {}).content || {}).parts
      ? data.candidates[0].content.parts.map(p => p.text || '').join('')
      : '';
    res.status(200).json({ text });
  } catch (e) {
    res.status(500).json({ error: String((e && e.message) || e) });
  }
};
