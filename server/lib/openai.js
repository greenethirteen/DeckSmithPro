import OpenAI from 'openai';

export function getOpenAIClient() {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('Missing OPENAI_API_KEY in server environment');
  return new OpenAI({ apiKey });
}

export function getModels() {
  return {
    text: process.env.OPENAI_TEXT_MODEL || 'gpt-4o',
    image: process.env.OPENAI_IMAGE_MODEL || 'gpt-image-1',
    // Gemini defaults (Nano Banana Pro for images)
    gemini_text: process.env.GEMINI_TEXT_MODEL || 'gemini-3-pro',
    gemini_image: process.env.GEMINI_IMAGE_MODEL || 'gemini-3-pro-image-preview'
  };
}
