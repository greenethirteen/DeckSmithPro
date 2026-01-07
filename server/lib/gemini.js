import 'dotenv/config';

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta';

function requireKey() {
  if (!GEMINI_API_KEY) {
    throw new Error('Missing GEMINI_API_KEY in server/.env (required for provider="gemini").');
  }
}

async function postGenerateContent(model, body) {
  requireKey();
  const url = `${GEMINI_API_BASE}/models/${encodeURIComponent(model)}:generateContent`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': GEMINI_API_KEY,
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Gemini API error (${res.status}): ${txt || res.statusText}`);
  }
  return res.json();
}

function extractText(responseJson) {
  const parts = responseJson?.candidates?.[0]?.content?.parts;
  if (!Array.isArray(parts)) return '';
  return parts.map(p => p?.text).filter(Boolean).join('\n');
}

function extractFirstInlineImage(responseJson) {
  const parts = responseJson?.candidates?.[0]?.content?.parts;
  if (!Array.isArray(parts)) return null;
  for (const p of parts) {
    // REST uses inlineData, older examples use inline_data
    const inline = p?.inlineData || p?.inline_data || p?.inlineData;
    if (inline?.data) {
      return {
        mimeType: inline.mimeType || inline.mime_type || 'image/png',
        base64: inline.data,
      };
    }
  }
  return null;
}

export async function geminiGenerateJson({ model, system, user, jsonSchema, temperature = 0.7 }) {
  // Use Structured Outputs: responseMimeType + responseJsonSchema
  const prompt = `${system}\n\n${user}`;
  const body = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature,
      responseMimeType: 'application/json',
      responseJsonSchema: jsonSchema,
    }
  };

  const out = await postGenerateContent(model, body);
  const txt = extractText(out);
  if (!txt) throw new Error('No JSON returned from Gemini.');
  try {
    return JSON.parse(txt);
  } catch {
    // Fallback: try to find the first JSON object in the text
    const m = txt.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    throw new Error('Gemini returned non-JSON output.');
  }
}

export async function geminiGenerateImage({ model, prompt, aspectRatio = '16:9', imageSize = '2K' }) {
  const body = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      // Ensures we actually get image parts back.
      responseModalities: ['Image'],
      imageConfig: {
        aspectRatio,
        ...(imageSize ? { imageSize } : {})
      }
    }
  };

  const out = await postGenerateContent(model, body);
  const img = extractFirstInlineImage(out);
  if (!img?.base64) {
    const txt = extractText(out);
    throw new Error(`Gemini did not return an image. ${txt ? `Text: ${txt}` : ''}`);
  }
  return img;
}
