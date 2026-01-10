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

  // Precompute prompts so we can report progress accurately.
  const prepared = (Array.isArray(slides) ? slides : []).map((slide, idx) => ({
    slide,
    idx,
    prompt: buildImagePrompt(slide, theme, opts)
  }));

  const total = prepared.filter(p => p.prompt).length;
  let done = 0;

  const results = await Promise.all(prepared.map((p) => {
    if (!p.prompt) return Promise.resolve({ idx: p.idx, file: null, prompt: null });

    return limit(async () => {
      let b64 = null;

      if (provider === 'gemini') {
        const img = await geminiGenerateImage({
          model: geminiImageModel,
          prompt: p.prompt,
          aspectRatio: opts.geminiAspectRatio || '16:9',
          imageSize: opts.geminiImageSize || '2K'
        });
        b64 = img.base64;
      } else {
        const result = await client.images.generate({
          model: openaiImageModel,
          prompt: p.prompt,
          // We'll resize/crop for PPTX anyway.
          size: opts.imageSize || '1024x1024'
        });
        b64 = result?.data?.[0]?.b64_json || null;
      }

      const rawName = `img_${p.idx}_${nanoid(6)}.png`;
      const rawPath = path.join(tmpDir, rawName);
      const full = path.join(tmpDir, `img_${p.idx}.jpg`);

      if (b64) {
        await fs.writeFile(rawPath, Buffer.from(b64, 'base64'));
        await coverTo(rawPath, full, 1600, 900);
        await fs.remove(rawPath).catch(() => {});
        return { idx: p.idx, file: full, prompt: p.prompt };
      }

      return { idx: p.idx, file: null, prompt: p.prompt };
    }).finally(() => {
      // Progress callback is optional; it drives UI feedback during image generation.
      if (p.prompt) {
        done += 1;
        try { opts.onProgress?.({ done, total, idx: p.idx, prompt: p.prompt }); } catch {}
      }
    });
  }));

  // Map by slide index
  const byIdx = {};
  for (const r of results) byIdx[r.idx] = r;
  return byIdx;
}

function buildImagePrompt(slide, theme, opts) {
  let base = (slide?.image_prompt || '').toString().trim();

  // If the planner omitted an image prompt for an image-heavy layout, create a safe fallback prompt.
  const layout = (slide?.layout || slide?.kind || '').toString().toLowerCase();
  const wantsImage = ['hero','full_bleed','split','section_header','image_caption','case_study','quote','execution_example','execution_examples']
    .some(k => layout.includes(k));

  if ((!base || /^none$/i.test(base)) && opts.autoImagePrompts !== false && wantsImage) {
    const t = (slide?.title || '').toString().trim();
    const sub = (slide?.subtitle || '').toString().trim();
    const b = Array.isArray(slide?.bullets) ? slide.bullets.slice(0, 4).join('; ') : '';
    base = [
      t ? `Background visual for: "${t}".` : 'Background visual for a presentation slide.',
      sub ? `Subtitle context: "${sub}".` : '',
      b ? `Key points: ${b}.` : '',
      'Use a metaphorical, business-safe scene that supports the message.'
    ].filter(Boolean).join(' ');
  }

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
