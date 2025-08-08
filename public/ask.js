// api/ask.js  (Node 18+ on Vercel)
import fetch from 'node-fetch';

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Only POST allowed' });
  const { prompt, tableData } = req.body || {};
  if (!prompt) return res.status(400).json({ error: 'Missing prompt' });

  const API_KEY = process.env.GEMINI_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Server missing GEMINI_API_KEY env' });

  try {
    // Shorten table data if too big, send first N rows only (adjust N)
    const maxRows = 200;
    const inputSample = (tableData && tableData.values) ? tableData.values.slice(0, maxRows) : [];

    // System instructions to get clean CSV/JSON results when needed
    const systemPrompt = `You are an expert spreadsheet assistant. Input is a JSON array-of-arrays (rows). The user prompt follows. If the user asks for a table, **output only CSV** (no explanation). If user asks for JSON, output only JSON (array-of-arrays or array-of-objects). Otherwise return plain text.`;

    // Build request body for Gemini-ish API. Replace if your provider differs.
    const body = {
      // Example for Google Generative Language (edit if your provider differs)
      prompt: {
        text: `${systemPrompt}\n\nTable sample (first ${inputSample.length} rows):\n${JSON.stringify(inputSample)}\n\nUser request:\n${prompt}`
      }
      // If your provider requires a different shape, replace above with required fields.
    };

    // Example endpoint (replace if your provider uses different URL)
    // Many Google examples use: https://generativelanguage.googleapis.com/v1beta2/models/{model}:generate
    const model = process.env.LLM_MODEL || 'models/text-bison-001'; // change if needed
    const endpoint = process.env.LLM_ENDPOINT || `https://generativelanguage.googleapis.com/v1beta2/${model}:generate?key=${API_KEY}`;

    const r = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });

    if (!r.ok) {
      const t = await r.text();
      console.error('LLM error', r.status, t);
      return res.status(502).json({ error: 'LLM error', details: t });
    }

    const parsed = await r.json();

    // Extract text candidate - adapt based on actual provider response
    // For Google: parsed.candidates[0].content[0].text or parsed.output[0].content[0].text
    let answer = '';
    if (parsed.candidates && parsed.candidates[0] && parsed.candidates[0].content) {
      answer = parsed.candidates[0].content.map(c => c.text || c).join('\n');
    } else if (parsed.output && parsed.output[0] && parsed.output[0].content) {
      // alternative shape
      answer = parsed.output[0].content.map(c => c.text || c).join('\n');
    } else {
      answer = JSON.stringify(parsed);
    }

    return res.json({ success: true, answer });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: err.message });
  }
}
