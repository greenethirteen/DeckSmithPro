import fs from 'fs-extra';
import path from 'path';
import { nanoid } from 'nanoid';
import pLimit from 'p-limit';
import { getOpenAIClient, getModels } from './openai.js';
import { geminiGenerateImage } from './gemini.js';
import { coverTo } from './util.js';
import { getDeckStylePreset } from './themes.js';

/**
 * Generate slide-supporting images and return file paths.
 * We generate square-ish images then crop/cover to slide needs.
 */
export async function generateImages(slides, theme, tmpDir, opts = {}) {
  const provider = (opts.provider || 'openai').toString().toLowerCase();
  const client = provider === 'openai' ? getOpenAIClient() : null;
  const { image: openaiImageModel, gemini_image: geminiImageModel } = getModels();
  const concurrency = Math.max(1, Math.min(3, Number(opts.imageConcurrency || 2)));
  const limit = pLimit(concurrency);

  const results = await Promise.all(slides.map((slide, idx) => limit(async () => {
    const prompt = buildImagePrompt(slide, theme, opts);
    if (!prompt) return { idx, file: null, prompt: null };

    let b64 = null;

    if (provider === 'gemini') {
      const img = await geminiGenerateImage({
        model: geminiImageModel,
        prompt,
        aspectRatio: opts.geminiAspectRatio || '16:9',
        imageSize: opts.geminiImageSize || '2K'
      });
      b64 = img.base64;
    } else {
      const result = await client.images.generate({
        model: openaiImageModel,
        prompt,
        // We'll resize/crop for PPTX anyway.
        size: opts.imageSize || '1024x1024'
      });
      b64 = result?.data?.[0]?.b64_json;
    }

    if (!b64) return { idx, file: null, prompt };

    const rawPath = path.join(tmpDir, `img_raw_${idx}_${nanoid(6)}.png`);
    await fs.writeFile(rawPath, Buffer.from(b64, 'base64'));

    // Create derivative (16:9) used by all renderers.
    const full = path.join(tmpDir, `img_full_${idx}_${nanoid(6)}.png`);
    await coverTo(rawPath, full, 1600, 900);

    return { idx, file: full, prompt };
  })));

  // Map by slide index
  const byIdx = {};
  for (const r of results) byIdx[r.idx] = r;
  return byIdx;
}

function buildImagePrompt(slide, theme, opts) {
  const base = (slide?.image_prompt || '').toString().trim();
  if (!base || /^none$/i.test(base)) return null;

  const preset = getDeckStylePreset(opts.deckStyle || opts.deck_style || theme.deck_style);
  const style = (opts.imageStyle || preset.imageHint || theme.vibe || 'Modern, premium').toString();
  const colorHint = `Primary color ${theme.primary_color}, secondary ${theme.secondary_color}`;
  const deckHint = preset?.promptHint ? `Deck visual style: ${preset.promptHint}.` : '';

  const constraints = [
    'No logos, no trademarks, no copyrighted characters.',
    'No text, no typography, no words, no headlines.',
    'High-quality, clean composition, generous negative space for overlaid text.',
    'If people are shown, keep faces non-identifiable and natural.',
    'Avoid random patterns unless explicitly asked; image must support the slide message.'
  ].join(' ');

  return `${base}\n\nStyle: ${style}. ${deckHint} ${colorHint}.\n${constraints}`;
}
