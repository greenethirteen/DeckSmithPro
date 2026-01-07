/**
 * Theme + style helpers.
 *
 * NOTE: These are "best-effort" style presets. PPTX will substitute fonts if missing.
 */

export function normalizeHex(hex, fallback = '#0B0F1A') {
  if (!hex) return fallback;
  const h = hex.toString().trim();
  if (/^#[0-9A-Fa-f]{6}$/.test(h)) return h.toUpperCase();
  return fallback;
}

export function pickFont(name, fallback) {
  const n = (name || '').toString().trim();
  if (!n) return fallback;
  return n;
}

/**
 * Deck style presets (renderer + prompting hints).
 * Keep this list small but clearly distinct.
 */
export const DECK_STYLE_PRESETS = {
  classic: {
    id: 'classic',
    label: 'Classic Premium',
    corner: 'rounded',
    headingFont: 'Helvetica',
    bodyFont: 'Futura',
    titleSize: 60,
    subtitleSize: 22,
    bodySize: 18,
    overlayLight: 35,
    overlayDark: 55,
    panelFill: 'FFFFFF',
    panelTransparency: 0,
    panelLine: 'FFFFFF',
    panelLineTransparency: 0,
    panelLineWidth: 1,
    borderWidth: 0,
    accentThickness: 0.08,
    promptHint: 'Premium editorial photography, clean gradients, high contrast, minimal, modern, no text',
    imageHint: 'Editorial photography + clean gradients, premium, high contrast'
  },

  neo_brutal: {
    id: 'neo_brutal',
    label: 'Neo‑Brutalist',
    corner: 'sharp',
    headingFont: 'Arial Black',
    bodyFont: 'Arial',
    titleSize: 74,
    subtitleSize: 22,
    bodySize: 18,
    overlayLight: 12,
    overlayDark: 28,
    panelFill: 'FFFFFF',
    panelTransparency: 0,
    panelLine: '000000',
    panelLineTransparency: 0,
    panelLineWidth: 3,
    borderWidth: 4,
    accentThickness: 0.16,
    promptHint: 'Neo-brutalist poster design, bold flat shapes, thick outlines, high contrast, slight grain, minimal palette, no text',
    imageHint: 'Neo‑brutalist poster, bold blocks, thick outlines, high contrast, slight grain'
  },

  bento_minimal: {
    id: 'bento_minimal',
    label: 'Bento Minimal',
    corner: 'rounded',
    headingFont: 'Inter',
    bodyFont: 'Inter',
    titleSize: 58,
    subtitleSize: 20,
    bodySize: 17,
    overlayLight: 18,
    overlayDark: 34,
    panelFill: 'FFFFFF',
    panelTransparency: 10,
    panelLine: 'FFFFFF',
    panelLineTransparency: 35,
    panelLineWidth: 1,
    borderWidth: 0,
    accentThickness: 0.08,
    promptHint: 'Minimal bento grid aesthetic, airy whitespace, soft shadows, product design vibe, clean lighting, no text',
    imageHint: 'Minimal bento grid, airy whitespace, soft shadows, product design aesthetic'
  },

  gradient_mesh: {
    id: 'gradient_mesh',
    label: 'Gradient Mesh',
    corner: 'rounded',
    headingFont: 'Helvetica Neue',
    bodyFont: 'Helvetica',
    titleSize: 66,
    subtitleSize: 22,
    bodySize: 18,
    overlayLight: 8,
    overlayDark: 22,
    panelFill: '0B0F1A',
    panelTransparency: 35,
    panelLine: '0B0F1A',
    panelLineTransparency: 100,
    panelLineWidth: 1,
    borderWidth: 0,
    accentThickness: 0.08,
    promptHint: 'Abstract gradient mesh, cinematic lighting, color blobs, glossy, soft noise, no text',
    imageHint: 'Abstract gradient mesh, cinematic, glossy color blobs, soft noise'
  }
};

export function getDeckStylePreset(id) {
  const key = (id || '').toString().trim().toLowerCase();
  if (DECK_STYLE_PRESETS[key]) return DECK_STYLE_PRESETS[key];
  return DECK_STYLE_PRESETS.classic;
}

export function listDeckStyles() {
  return Object.values(DECK_STYLE_PRESETS).map(s => ({ id: s.id, label: s.label }));
}
