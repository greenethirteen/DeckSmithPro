import path from 'path';
import fs from 'fs-extra';
import PptxGenJS from 'pptxgenjs';
import sharp from 'sharp';
import { spawn } from 'child_process';
import { generateImages } from './images.js';
import { ensureCleanTmp, getLuminance } from './util.js';
import { normalizeHex, pickFont, getDeckStylePreset } from './themes.js';

const SLIDE_W = 13.333; // inches for LAYOUT_WIDE
const SLIDE_H = 7.5;

async function spawnCmd(cmd, args, opts = {}) {
  return await new Promise((resolve, reject) => {
    const p = spawn(cmd, args, { ...opts, stdio: ['ignore', 'ignore', 'pipe'] });
    let stderr = '';
    p.stderr?.on('data', (d) => { stderr += d.toString(); });
    p.on('error', reject);
    p.on('close', (code) => {
      if (code === 0) return resolve();
      reject(new Error(`${cmd} exited with code ${code}. ${stderr}`.trim()));
    });
  });
}

async function runSofficeConvert({ sofficePath = 'soffice', pptxPath, outDir }) {
  await fs.ensureDir(outDir);
  // LibreOffice writes one PNG per slide.
  await spawnCmd(sofficePath, [
    '--headless',
    '--nologo',
    '--nofirststartwizard',
    '--convert-to', 'png',
    '--outdir', outDir,
    pptxPath
  ], { env: process.env });
}

async function findConvertedPngForIndex(outDir, baseName, index) {
  const files = await fs.readdir(outDir).catch(() => []);
  const pngs = files.filter(f => f.toLowerCase().endsWith('.png') && f.startsWith(baseName));
  if (!pngs.length) return null;

  const parseIdx = (f) => {
    const lower = f.toLowerCase();
    const baseLower = baseName.toLowerCase();
    if (lower === `${baseLower}.png`) return 0;
    let m = lower.match(new RegExp(`^${baseLower}_(\\d+)\\.png$`));
    if (m) return Number(m[1]);
    m = lower.match(new RegExp(`^${baseLower}-(\\d+)\\.png$`));
    if (m) return Number(m[1]);
    m = lower.match(new RegExp(`^${baseLower}\\s*\\((\\d+)\\)\\.png$`));
    if (m) return Number(m[1]);
    return null;
  };

  for (const f of pngs) {
    const idx = parseIdx(f);
    if (idx === index) return path.join(outDir, f);
  }

  // Fallback: newest file (still gives a "live" feel)
  const stats = await Promise.all(pngs.map(async f => ({ f, st: await fs.stat(path.join(outDir, f)).catch(() => null) })));
  const newest = stats.filter(x => x.st).sort((a,b)=>b.st.mtimeMs - a.st.mtimeMs)[0];
  return newest ? path.join(outDir, newest.f) : null;
}


function overlayTransparency(theme, luminance) {
  const st = theme.style || {};
  const light = Number.isFinite(+st.overlayLight) ? +st.overlayLight : 35;
  const dark = Number.isFinite(+st.overlayDark) ? +st.overlayDark : 55;
  return luminance > 0.55 ? light : dark;
}

function panelShapeType(pptx, theme) {
  return (theme.style?.corner === 'sharp') ? pptx.ShapeType.rect : pptx.ShapeType.roundRect;
}

function panelStyle(theme, overrides = {}) {
  const st = theme.style || {};
  return {
    fill: { color: overrides.fillColor || st.panelFill || 'FFFFFF', transparency: overrides.fillTransparency ?? st.panelTransparency ?? 0 },
    line: { color: overrides.lineColor || st.panelLine || (st.panelFill || 'FFFFFF'), transparency: overrides.lineTransparency ?? st.panelLineTransparency ?? 0, width: overrides.lineWidth ?? st.panelLineWidth ?? 1 }
  };
}

function svgToDataUri(svg) {
  return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
}

function infographicIconSvg(kind, color) {
  const stroke = color || '#F04C2E';
  if (kind === 'signal') {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="120" height="120" viewBox="0 0 120 120" fill="none" stroke="${stroke}" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"><circle cx="60" cy="60" r="10"/><circle cx="60" cy="60" r="26"/><circle cx="60" cy="60" r="42"/></svg>`;
  }
  if (kind === 'hashtag') {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="120" height="120" viewBox="0 0 120 120" fill="none" stroke="${stroke}" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"><line x1="40" y1="20" x2="32" y2="100"/><line x1="78" y1="20" x2="70" y2="100"/><line x1="20" y1="44" x2="100" y2="44"/><line x1="16" y1="76" x2="96" y2="76"/><circle cx="92" cy="92" r="16"/></svg>`;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="120" height="120" viewBox="0 0 120 120" fill="none" stroke="${stroke}" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"><path d="M60 110C60 110 32 76 32 50C32 33 45 20 60 20C75 20 88 33 88 50C88 76 60 110 60 110Z"/><circle cx="60" cy="50" r="12"/></svg>`;
}

function addSlideFrame(slide, pptx, theme) {
  const bw = Number.isFinite(+theme.style?.borderWidth) ? +theme.style.borderWidth : 0;
  if (bw <= 0) return;
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.12, y: 0.12, w: SLIDE_W - 0.24, h: SLIDE_H - 0.24,
    fill: { color: 'FFFFFF', transparency: 100 },
    line: { color: theme.secondary, width: bw }
  });
}

async function resolveBrandLogo(plan, options, tmpDir) {
  const raw = (options.brandLogo || plan.brand_logo || plan.theme?.brand_logo || '').toString().trim();
  if (!raw) return null;
  if (raw.startsWith('data:image/')) return { data: raw };
  if (/^https?:\/\//i.test(raw)) {
    try {
      const res = await fetch(raw);
      if (!res.ok) return null;
      const buf = Buffer.from(await res.arrayBuffer());
      const outPath = path.join(tmpDir, `brand_logo_${Date.now()}.png`);
      await fs.writeFile(outPath, buf);
      return { path: outPath };
    } catch {
      return null;
    }
  }
  if (await fs.pathExists(raw)) return { path: raw };
  return null;
}

function addBrandLogoToSlide(slide, logo) {
  if (!logo || !slide) return;
  const logoH = 0.35;
  const logoW = 1.4;
  slide.addImage({
    ...(logo.path ? { path: logo.path } : { data: logo.data }),
    x: SLIDE_W - logoW - 0.5,
    y: 0.35,
    w: logoW,
    h: logoH
  });
}

// Cache for resized/cropped image variants so we never stretch images.
// Keyed by `${src}|${w.toFixed(3)}x${h.toFixed(3)}`.
async function coverImage(srcPath, targetWIn, targetHIn, tmpDir, cache) {
  if (!srcPath) return null;
  const key = `${srcPath}|${targetWIn.toFixed(3)}x${targetHIn.toFixed(3)}`;
  if (cache.has(key)) return cache.get(key);

  // We render at a reasonable pixel density so exports are crisp but not huge.
  // Use height as the anchor for most placements.
  const aspect = targetWIn / targetHIn;
  const outH = 1080; // px
  const outW = Math.max(320, Math.round(outH * aspect));
  const outPath = path.join(tmpDir, `cover_${cache.size}_${outW}x${outH}.jpg`);

  await sharp(srcPath)
    .resize(outW, outH, { fit: 'cover', position: 'attention' })
    .jpeg({ quality: 88 })
    .toFile(outPath);

  cache.set(key, outPath);
  return outPath;
}

function headlineStyle(theme, fontSize) {
  return {
    fontFace: theme.headingFont,
    fontSize,
    bold: true,
    fit: 'shrink' // prevent overflow in narrow boxes
  };
}

function bodyStyle(theme, fontSize) {
  return {
    fontFace: theme.bodyFont,
    fontSize,
    fit: 'shrink'
  };
}

export async function exportPptx(plan, options = {}, ctx = {}) {
  const tmpDir = ctx.tmpDir || path.resolve('.tmp');
  await ensureCleanTmp(tmpDir);

  const imgCropCache = new Map();

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';

  // Theme + deck style preset
  const deckStyleId = options.deckStyle || options.deck_style || plan.theme?.deck_style;
  const preset = getDeckStylePreset(deckStyleId);

  const theme = {
    primary: normalizeHex(plan.theme?.primary_color, '#0B0F1A'),
    secondary: normalizeHex(plan.theme?.secondary_color, '#2A7FFF'),
    // Fonts fall back to preset if plan theme fonts are missing.
    headingFont: pickFont(plan.theme?.font_heading, preset.headingFont),
    bodyFont: pickFont(plan.theme?.font_body, preset.bodyFont),
    vibe: plan.theme?.vibe || 'Modern, premium',
    style: preset
  };

  pptx.author = 'DeckSmith Pro';
  pptx.company = '';
  pptx.subject = plan.deck_title || 'Deck';
  pptx.title = plan.deck_title || 'Deck';

  ctx.onStatus?.({ phase: 'generating_images', message: 'Generating images…' });
  const imgMap = await generateImages(plan.slides, {
    vibe: theme.vibe,
    primary_color: theme.primary,
    secondary_color: theme.secondary
  }, tmpDir, {
    provider: options.provider || 'openai',
    imageConcurrency: options.imageConcurrency ?? 2,
    imageSize: options.imageSize ?? "1024x1024",
    imageStyle: options.imageStyle ?? theme.style?.imageHint ?? theme.vibe,
    deckStyle: theme.style?.id
  });

  const brandLogo = await resolveBrandLogo(plan, options, tmpDir);

  ctx.onStatus?.({ phase: 'images_ready', message: 'Images ready. Rendering slides…' });

  // Slide rendering
  for (let i = 0; i < plan.slides.length; i++) {
    const s = plan.slides[i];
    ctx.onSlide?.({ index: i, total: plan.slides.length, slide: {
      kind: s?.kind || s?.type || '',
      layout: s?.layout || (i === 0 ? 'hero' : 'split'),
      title: s?.title || '',
      subtitle: s?.subtitle || '',
      bullets: Array.isArray(s?.bullets) ? s.bullets : [],
      section: s?.section || ''
    }});

    const imageFile = imgMap?.[i]?.file || null;
    await renderByLayout(pptx, s, plan, theme, imageFile, i, tmpDir, imgCropCache);
    const renderedSlide = pptx._slides?.[pptx._slides.length - 1];
    if (renderedSlide && brandLogo) {
      addBrandLogoToSlide(renderedSlide, brandLogo);
    }

    // Optional: true live thumbnails via LibreOffice conversion.
    // NOTE: This is intentionally "easy mode" (writes partial PPTX + converts after each slide).
    if (ctx?.sofficeThumbs && ctx?.jobId && ctx?.thumbDir && typeof ctx?.onThumbnail === 'function') {
      try {
        const partialPptx = path.join(tmpDir, `partial_${ctx.jobId}.pptx`);
        await pptx.writeFile({ fileName: partialPptx });
        await runSofficeConvert({
          sofficePath: ctx.sofficePath || process.env.SOFFICE_PATH || 'soffice',
          pptxPath: partialPptx,
          outDir: ctx.thumbDir
        });
        const base = path.basename(partialPptx, path.extname(partialPptx));
        const pngPath = await findConvertedPngForIndex(ctx.thumbDir, base, i);
        if (pngPath) {
          ctx.onThumbnail({ index: i, total: plan.slides.length, path: pngPath });
        }
      } catch (e) {
        // If LibreOffice isn't available, we silently skip thumbnails (export still succeeds).
        ctx.onStatus?.({ phase: 'thumbs_skipped', message: 'Thumbnail rendering skipped (LibreOffice not available).' });
      }
    }
  }

  ctx.onStatus?.({ phase: 'finalizing', message: 'Finalizing PPTX…' });

  // Write to buffer
  const buf = await pptx.write('nodebuffer');
  ctx.onStatus?.({ phase: 'done', message: 'PPTX ready.' });
  return buf;
}

async function renderByLayout(pptx, slide, plan, theme, imageFile, idx, tmpDir, imgCropCache) {
  const layout = slide.layout || (idx === 0 ? 'hero' : 'split');
  switch (layout) {
    case 'hero':
      return renderHero(pptx, slide, plan, theme, imageFile, tmpDir, imgCropCache);
    case 'agency_center':
      return renderAgencyCenter(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'agency_half':
      return renderAgencyHalf(pptx, slide, theme, imageFile, idx, tmpDir, imgCropCache);
    case 'agency_infographic':
      return renderAgencyInfographic(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'section_header':
      return renderSectionHeader(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'agenda':
      return renderAgenda(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'cards':
      return renderCards(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'image_caption':
      return renderImageCaption(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'timeline':
      return renderTimeline(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'kpi_dashboard':
      return renderKpiDashboard(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'traffic_light':
      return renderTrafficLight(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'table':
      return renderTable(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'pricing':
      return renderPricing(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'comparison_matrix':
      return renderComparisonMatrix(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'process_steps':
      return renderProcessSteps(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'team_grid':
      return renderTeamGrid(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'logo_wall':
      return renderLogoWall(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'cta':
      return renderCTA(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'swot':
      return renderSWOT(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'funnel':
      return renderFunnel(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'now_next_later':
      return renderNowNextLater(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'okr':
      return renderOKR(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'case_study':
      return renderCaseStudy(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'chart_bar':
      return renderChartBar(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'chart_line':
      return renderChartLine(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'org_chart':
      return renderOrgChart(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'faq':
      return renderFAQ(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'appendix':
      return renderAppendix(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'infographic_3':
      return renderInfographic3(pptx, slide, theme);
    case 'full_bleed':
      return renderFullBleed(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'quote':
      return renderQuote(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'stats':
      return renderStats(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'two_column':
      return renderTwoColumn(pptx, slide, theme, imageFile, tmpDir, imgCropCache);
    case 'split':
    default:
      return renderSplit(pptx, slide, theme, imageFile, idx, tmpDir, imgCropCache);
  }
}

// ---------- Layouts ----------
async function renderHero(pptx, slide, plan, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme);
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    const L = await getLuminance(bg);
    // Overlay for readability
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: theme.primary, transparency: overlayTransparency(theme, L) },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.background = { color: theme.primary };
  }

  // Big title
  // IMPORTANT: only the *cover* slide should inherit plan.deck_title / plan.deck_subtitle.
  // For hero-style slides like `big_idea` / `creative_concept`, we must use the slide's own title.
  const isCoverLike = ['cover', 'title', 'title_page'].includes(slide.kind) || (slide.section === 'title');
  const title = (isCoverLike && plan.deck_title) ? plan.deck_title : (slide.title || plan.deck_title || '');
  const subtitle = (isCoverLike && plan.deck_subtitle) ? plan.deck_subtitle : (slide.subtitle || '');

  s.addText(title, {
    x: 0.9, y: 2.2, w: SLIDE_W - 1.8, h: 2.0,
    ...headlineStyle(theme, theme.style?.titleSize ?? 60),
    color: 'FFFFFF',
  });

  if (subtitle) {
    s.addText(subtitle, {
      x: 0.9, y: 4.55, w: SLIDE_W - 1.8, h: 1.0,
      ...bodyStyle(theme, theme.style?.subtitleSize ?? 20),
      color: 'FFFFFF',
    });
  }

  // Accent bar
  s.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 1.85, w: 1.6, h: 0.08,
    fill: { color: theme.secondary },
    line: { color: theme.secondary }
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderSplit(pptx, slide, theme, imageFile, idx, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme);

  // Left visual
  if (imageFile) {
    const bg = await coverImage(imageFile, 6.1, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: 6.1, h: SLIDE_H });
  }
  else {
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: 6.1, h: SLIDE_H,
      fill: { color: theme.primary },
      line: { color: theme.primary }
    });
  }
  // Subtle overlay
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 6.1, h: SLIDE_H,
    fill: { color: theme.primary, transparency: 70 },
    line: { color: theme.primary, transparency: 100 }
  });

  // Right content panel
  s.addShape(panelType, {
    x: 6.0, y: 0.6, w: 7.0, h: 6.3,
    ...pStyle
  });

  // Accent notch
  s.addShape(pptx.ShapeType.roundRect, {
    x: 6.0, y: 0.6, w: 0.22, h: 6.3,
    fill: { color: theme.secondary },
    line: { color: theme.secondary }
  });

  // Typography variation: alternating scale
  const big = idx % 2 === 0;
  const titleSize = big ? 40 : 36;

  s.addText(slide.title, {
    x: 6.35, y: 1.0, w: 6.4, h: 0.9,
    ...headlineStyle(theme, titleSize),
    color: theme.primary
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 6.35, y: 1.85, w: 6.4, h: 0.6,
      ...bodyStyle(theme, 16),
      color: '3B3B3B'
    });
  }

  // Bullets
  const bulletText = (slide.bullets || []).filter(Boolean).map(b => `• ${b}`).join('\n');
  if (bulletText) {
    s.addText(bulletText, {
      x: 6.35, y: 2.55, w: 6.4, h: 3.4,
      ...bodyStyle(theme, theme.style?.bodySize ?? 18),
      color: '111111',
      valign: 'top',
      lineSpacingMultiple: 1.15
    });
  }

  // Footer label (tiny)
  s.addText(`Slide ${idx + 1}`, {
    x: 6.35, y: 6.4, w: 6.4, h: 0.3,
    ...bodyStyle(theme, 10),
    color: '9A9A9A'
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderFullBleed(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: theme.primary, transparency: L > 0.55 ? 40 : 65 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.background = { color: theme.primary };
  }

  // Big headline centered
  s.addText(slide.title, {
    x: 1.2, y: 2.4, w: SLIDE_W - 2.4, h: 1.4,
    ...headlineStyle(theme, 56),
    color: 'FFFFFF',
    align: 'center'
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 1.8, y: 3.9, w: SLIDE_W - 3.6, h: 0.8,
      ...bodyStyle(theme, theme.style?.bodySize ?? 18),
      color: 'FFFFFF',
      align: 'center'
    });
  }

  // Bullet strip at bottom
  const bullets = (slide.bullets || []).slice(0, 3);
  if (bullets.length) {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 1.2, y: 6.0, w: SLIDE_W - 2.4, h: 1.1,
      fill: { color: 'FFFFFF', transparency: 15 },
      line: { color: 'FFFFFF', transparency: 100 }
    });
    s.addText(bullets.map(b => `• ${b}`).join('   '), {
      x: 1.4, y: 6.15, w: SLIDE_W - 2.8, h: 0.8,
      ...bodyStyle(theme, 16),
      color: theme.primary,
      align: 'center'
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderQuote(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme);

  // Background
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: theme.primary, transparency: 70 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.background = { color: theme.primary };
  }

  const q = slide.quote?.text || slide.subtitle || (slide.bullets?.[0] ?? '');
  const by = slide.quote?.attribution || '';

  // Big quote marks
  s.addText('“', {
    x: 0.9, y: 1.2, w: 1, h: 1,
    ...headlineStyle(theme, 96),
    color: theme.secondary
  });

  s.addText(q, {
    x: 1.6, y: 1.8, w: SLIDE_W - 3.2, h: 3.5,
    ...headlineStyle(theme, 44),
    color: 'FFFFFF'
  });

  if (by) {
    s.addText(`— ${by}`, {
      x: 1.6, y: 5.6, w: SLIDE_W - 3.2, h: 0.6,
      ...bodyStyle(theme, 16),
      color: 'FFFFFF'
    });
  }

  // Title as a small tag
  s.addShape(pptx.ShapeType.roundRect, {
    x: 1.6, y: 6.45, w: 4.2, h: 0.5,
    fill: { color: 'FFFFFF', transparency: 15 },
    line: { color: 'FFFFFF', transparency: 100 }
  });
  s.addText(slide.title, {
    x: 1.8, y: 6.53, w: 3.8, h: 0.4,
    ...bodyStyle(theme, 12),
    color: theme.primary
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderInfographic3(pptx, slide, theme) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title || 'Three Pillars', {
    x: 0.9, y: 0.6, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 36),
    color: '111111',
    align: 'center'
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 1.6, y: 1.45, w: SLIDE_W - 3.2, h: 0.6,
      ...bodyStyle(theme, 16),
      color: '444444',
      align: 'center'
    });
  }

  const cards = safeArr(slide.cards).slice(0, 3);
  const items = cards.length ? cards : [
    { title: 'Pillar One', body: 'Description.', tag: 'pin' },
    { title: 'Pillar Two', body: 'Description.', tag: 'signal' },
    { title: 'Pillar Three', body: 'Description.', tag: 'hashtag' }
  ];

  const colW = (SLIDE_W - 2.4) / 3;
  const iconSize = 1.6;
  const iconY = 2.6;

  for (let i = 0; i < 3; i++) {
    const x = 0.8 + i * colW;
    const iconX = x + (colW - iconSize) / 2;
    const iconKind = (items[i]?.tag || '').toString().toLowerCase();
    const svg = infographicIconSvg(iconKind, theme.secondary);
    s.addImage({ data: svgToDataUri(svg), x: iconX, y: iconY, w: iconSize, h: iconSize });

    s.addText([
      { text: `${i + 1}. `, options: { color: theme.secondary } },
      { text: items[i]?.title || `Pillar ${i + 1}` }
    ], {
      x, y: 4.35, w: colW - 0.2, h: 0.5,
      ...headlineStyle(theme, 16),
      color: '111111',
      align: 'center'
    });

    s.addText(items[i]?.body || '—', {
      x, y: 4.9, w: colW - 0.2, h: 1.4,
      ...bodyStyle(theme, 12),
      color: '444444',
      align: 'center'
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderStats(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme);

  // Background split: left solid, right image
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 6.0, h: SLIDE_H,
    fill: { color: theme.primary },
    line: { color: theme.primary }
  });
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W - 6.0, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 6.0, y: 0, w: SLIDE_W - 6.0, h: SLIDE_H });
    s.addShape(pptx.ShapeType.rect, {
      x: 6.0, y: 0, w: SLIDE_W - 6.0, h: SLIDE_H,
      fill: { color: theme.primary, transparency: 65 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.addShape(pptx.ShapeType.rect, {
      x: 6.0, y: 0, w: SLIDE_W - 6.0, h: SLIDE_H,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });
  }

  const statVal = slide.stat?.value || (slide.bullets?.[0] ?? '');
  const statLabel = slide.stat?.label || slide.subtitle || '';

  s.addText(slide.title, {
    x: 0.8, y: 0.8, w: 5.0, h: 0.8,
    ...bodyStyle(theme, 14),
    color: 'FFFFFF'
  });

  s.addText(statVal, {
    x: 0.8, y: 2.1, w: 5.0, h: 2.2,
    ...headlineStyle(theme, 78),
    color: theme.secondary
  });

  if (statLabel) {
    s.addText(statLabel, {
      x: 0.8, y: 4.25, w: 5.0, h: 0.9,
      ...bodyStyle(theme, theme.style?.bodySize ?? 18),
      color: 'FFFFFF'
    });
  }

  // Supporting bullets on right
  const bullets = (slide.bullets || []).slice(0, 4).map(b => `• ${b}`).join('\n');
  if (bullets) {
    s.addText(bullets, {
      x: 6.4, y: 1.4, w: SLIDE_W - 6.9, h: 5.6,
      ...bodyStyle(theme, theme.style?.bodySize ?? 18),
      color: 'FFFFFF',
      lineSpacingMultiple: 1.12
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderTwoColumn(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme);
  // Light background
  s.background = { color: 'FFFFFF' };

  // Top header band with optional image strip
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, 2.2, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: 2.2 });
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 2.2,
      fill: { color: theme.primary, transparency: 70 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 2.2,
      fill: { color: theme.primary },
      line: { color: theme.primary }
    });
  }

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.9,
    ...headlineStyle(theme, 46),
    color: 'FFFFFF'
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.55, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: 'FFFFFF'
    });
  }

  // Two columns
  const bullets = (slide.bullets || []).filter(Boolean);
  const mid = Math.ceil(bullets.length / 2);
  const left = bullets.slice(0, mid);
  const right = bullets.slice(mid);

  const box = (x, y, w, h, title, lines) => {
    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w, h,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });
    s.addText(title, {
      x: x + 0.35, y: y + 0.25, w: w - 0.7, h: 0.4,
      ...bodyStyle(theme, 12),
      color: '666666'
    });
    s.addShape(pptx.ShapeType.rect, {
      x: x + 0.35, y: y + 0.65, w: 0.9, h: 0.07,
      fill: { color: theme.secondary }, line: { color: theme.secondary }
    });
    s.addText(lines.map(t => `• ${t}`).join('\n'), {
      x: x + 0.35, y: y + 0.9, w: w - 0.7, h: h - 1.1,
      ...bodyStyle(theme, theme.style?.bodySize ?? 18),
      color: theme.primary,
      valign: 'top', lineSpacingMultiple: 1.12
    });
  };

  box(0.9, 2.6, 5.9, 4.5, 'Key points', left);
  box(6.55, 2.6, 5.9, 4.5, 'More detail', right.length ? right : ['—']);

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

// ---------- Agency / creative layouts (bold typographic) ----------

async function renderAgencyCenter(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  // Background: optional image, washed with a white overlay for clean type.
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: 'FFFFFF', transparency: 68 },
      line: { color: 'FFFFFF', transparency: 100 }
    });
  } else {
    s.background = { color: 'FFFFFF' };
  }

  const title = (slide.title || '').toString().trim();
  const subtitle = (slide.subtitle || '').toString().trim();
  const support = (slide.bullets || []).filter(Boolean).slice(0, 2);

  // Accent rule
  s.addShape(pptx.ShapeType.rect, {
    x: 6.0, y: 1.65, w: 1.33, h: 0.08,
    fill: { color: theme.secondary }, line: { color: theme.secondary }
  });

  s.addText(title || '—', {
    x: 1.0, y: 2.05, w: SLIDE_W - 2.0, h: 2.0,
    fontFace: theme.headingFont,
    fontSize: Math.min(86, (theme.style?.titleSize ?? 78) + 8),
    bold: true,
    color: theme.primary,
    align: 'center',
    valign: 'mid',
    lineSpacingMultiple: 0.92
  });

  if (subtitle) {
    s.addText(subtitle, {
      x: 2.0, y: 4.35, w: SLIDE_W - 4.0, h: 0.65,
      ...bodyStyle(theme, theme.style?.subtitleSize ?? 22),
      color: theme.primary,
      align: 'center',
      valign: 'mid'
    });
  }

  if (support.length) {
    s.addText(support.join('  •  '), {
      x: 2.0, y: 5.05, w: SLIDE_W - 4.0, h: 0.55,
      ...bodyStyle(theme, 16),
      color: '555555',
      align: 'center'
    });
  }

  // Tiny section label (optional)
  const section = (slide.section || '').toString().trim();
  if (section) {
    s.addText(section.toUpperCase(), {
      x: 0.9, y: 6.95, w: SLIDE_W - 1.8, h: 0.3,
      ...bodyStyle(theme, 10),
      color: '777777',
      align: 'center'
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderAgencyHalf(pptx, slide, theme, imageFile, idx, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  const imageLeft = (idx % 2 === 0);
  const imgX = imageLeft ? 0 : SLIDE_W / 2;
  const txtX = imageLeft ? SLIDE_W / 2 : 0;

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W / 2, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: imgX, y: 0, w: SLIDE_W / 2, h: SLIDE_H });
    // Subtle wash for polish
    s.addShape(pptx.ShapeType.rect, {
      x: imgX, y: 0, w: SLIDE_W / 2, h: SLIDE_H,
      fill: { color: 'FFFFFF', transparency: 18 },
      line: { color: 'FFFFFF', transparency: 100 }
    });
  } else {
    // If image missing, give the left/right half a gentle tint so it doesn't look broken.
    s.addShape(pptx.ShapeType.rect, {
      x: imgX, y: 0, w: SLIDE_W / 2, h: SLIDE_H,
      fill: { color: 'F3F4F6' },
      line: { color: 'F3F4F6' }
    });
  }

  // Text panel
  s.addShape(pptx.ShapeType.rect, {
    x: txtX, y: 0, w: SLIDE_W / 2, h: SLIDE_H,
    fill: { color: 'FFFFFF' },
    line: { color: 'FFFFFF', transparency: 100 }
  });

  // Accent bar
  s.addShape(pptx.ShapeType.rect, {
    x: txtX + 0.9, y: 1.05, w: 1.1, h: 0.08,
    fill: { color: theme.secondary }, line: { color: theme.secondary }
  });

  const title = (slide.title || '').toString().trim();
  const subtitle = (slide.subtitle || '').toString().trim();
  const bullets = (slide.bullets || []).filter(Boolean).slice(0, 2);

  // Centered within the text half
  s.addText(title || '—', {
    x: txtX + 0.9, y: 1.35, w: (SLIDE_W / 2) - 1.8, h: 2.2,
    fontFace: theme.headingFont,
    fontSize: Math.min(72, theme.style?.titleSize ?? 78),
    bold: true,
    color: theme.primary,
    align: 'center',
    valign: 'mid',
    lineSpacingMultiple: 0.95
  });

  if (subtitle) {
    s.addText(subtitle, {
      x: txtX + 1.2, y: 3.55, w: (SLIDE_W / 2) - 2.4, h: 0.85,
      ...bodyStyle(theme, theme.style?.subtitleSize ?? 22),
      color: '333333',
      align: 'center'
    });
  }

  if (bullets.length) {
    s.addText(bullets.map(b => `• ${b}`).join('\n'), {
      x: txtX + 1.25, y: 4.45, w: (SLIDE_W / 2) - 2.5, h: 2.5,
      ...bodyStyle(theme, 16),
      color: '555555',
      align: 'center',
      valign: 'top',
      lineSpacingMultiple: 1.12
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderAgencyInfographic(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  // Optional soft background texture
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: 'FFFFFF', transparency: 82 },
      line: { color: 'FFFFFF', transparency: 100 }
    });
  }

  const title = (slide.title || '').toString().trim();
  const subtitle = (slide.subtitle || '').toString().trim();
  s.addText(title || '—', {
    x: 0.9, y: 0.8, w: SLIDE_W - 1.8, h: 0.9,
    fontFace: theme.headingFont,
    fontSize: 50,
    bold: true,
    color: theme.primary
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.9, y: 1.65, w: SLIDE_W - 1.8, h: 0.55,
      ...bodyStyle(theme, 18),
      color: '333333'
    });
  }

  // Build 3 items from cards (preferred) or bullets.
  const cards = Array.isArray(slide.cards) ? slide.cards : null;
  const bullets = (slide.bullets || []).filter(Boolean);
  const items = (cards && cards.length)
    ? cards.slice(0, 3).map(c => ({ h: c.title, b: c.body || c.tag || '' }))
    : bullets.slice(0, 3).map((b, i) => ({ h: `Point ${i + 1}`, b }));

  while (items.length < 3) items.push({ h: `Point ${items.length + 1}`, b: '—' });

  const colW = (SLIDE_W - 2.2) / 3;
  const y = 2.55;
  for (let i = 0; i < 3; i++) {
    const x = 0.75 + i * colW;

    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: colW - 0.25, h: 4.45,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });

    // Big marker
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.35, y: y + 0.35, w: 0.72, h: 0.55,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });
    s.addText(String(i + 1), {
      x: x + 0.35, y: y + 0.35, w: 0.72, h: 0.55,
      fontFace: theme.headingFont,
      fontSize: 18,
      bold: true,
      color: 'FFFFFF',
      align: 'center',
      valign: 'mid'
    });

    s.addText(items[i].h || `Point ${i + 1}`, {
      x: x + 0.35, y: y + 1.05, w: colW - 0.95, h: 0.75,
      fontFace: theme.headingFont,
      fontSize: 20,
      bold: true,
      color: theme.primary,
      align: 'left'
    });

    s.addText(items[i].b || '—', {
      x: x + 0.35, y: y + 1.75, w: colW - 0.95, h: 3.0,
      ...bodyStyle(theme, 16),
      color: '444444',
      valign: 'top',
      lineSpacingMultiple: 1.12
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

// ---------- New reusable business layouts ----------

async function renderSectionHeader(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: theme.primary, transparency: L > 0.55 ? 30 : 60 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.background = { color: theme.primary };
  }

  // Small tag
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.9, y: 0.85, w: 2.0, h: 0.45,
    fill: { color: 'FFFFFF', transparency: 15 },
    line: { color: 'FFFFFF', transparency: 100 }
  });
  s.addText((slide.kind || 'Section').toString().toUpperCase(), {
    x: 1.05, y: 0.95, w: 1.7, h: 0.3,
    ...bodyStyle(theme, 11),
    color: 'FFFFFF'
  });

  s.addText(slide.title, {
    x: 0.9, y: 2.2, w: SLIDE_W - 1.8, h: 1.8,
    ...headlineStyle(theme, 68),
    color: 'FFFFFF'
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 4.2, w: SLIDE_W - 1.8, h: 1.2,
      ...bodyStyle(theme, 20),
      color: 'FFFFFF'
    });
  }

  s.addShape(pptx.ShapeType.rect, {
    x: 0.9, y: 1.95, w: 2.2, h: 0.08,
    fill: { color: theme.secondary },
    line: { color: theme.secondary }
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderAgenda(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  // Optional top strip image
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, 1.6, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: 1.6 });
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 1.6,
      fill: { color: theme.primary, transparency: 70 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: 1.2,
      fill: { color: theme.primary },
      line: { color: theme.primary }
    });
  }

  s.addText(slide.title || 'Agenda', {
    x: 0.9, y: 0.35, w: SLIDE_W - 1.8, h: 0.7,
    ...headlineStyle(theme, 44),
    color: 'FFFFFF'
  });

  const items = (slide.agenda_items || slide.bullets || []).filter(Boolean).slice(0, 8);
  const startY = 1.85;
  const rowH = 0.65;

  items.forEach((it, i) => {
    const y = startY + i * rowH;
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.9, y, w: 0.5, h: 0.45,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });
    s.addText(String(i + 1), {
      x: 0.9, y: y + 0.08, w: 0.5, h: 0.3,
      ...bodyStyle(theme, 14),
      color: 'FFFFFF',
      align: 'center'
    });
    s.addText(it, {
      x: 1.55, y: y + 0.02, w: SLIDE_W - 2.45, h: 0.5,
      ...bodyStyle(theme, 20),
      color: theme.primary
    });
  });

  // Subtle sidebar
  s.addShape(pptx.ShapeType.rect, {
    x: SLIDE_W - 0.3, y: 1.5, w: 0.18, h: SLIDE_H - 1.7,
    fill: { color: theme.secondary, transparency: 15 },
    line: { color: theme.secondary, transparency: 100 }
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderCards(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  // Header
  s.addText(slide.title, {
    x: 0.9, y: 0.7, w: SLIDE_W - 1.8, h: 0.9,
    ...headlineStyle(theme, 46),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.55, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const cards = Array.isArray(slide.cards) && slide.cards.length
    ? slide.cards
    : (slide.bullets || []).slice(0, 4).map((b, i) => ({ title: `Point ${i + 1}`, body: b, tag: '' }));

  const cols = cards.length <= 3 ? 3 : 4;
  const gap = 0.35;
  const cardW = (SLIDE_W - 1.8 - gap * (cols - 1)) / cols;
  const cardH = 4.6;
  const topY = 2.35;

  cards.slice(0, cols).forEach((c, i) => {
    const x = 0.9 + i * (cardW + gap);
    s.addShape(pptx.ShapeType.roundRect, {
      x, y: topY, w: cardW, h: cardH,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });

    // Accent chip
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.35, y: topY + 0.35, w: 0.65, h: 0.3,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });

    s.addText(c.tag || '', {
      x: x + 1.1, y: topY + 0.32, w: cardW - 1.45, h: 0.35,
      ...bodyStyle(theme, 11),
      color: '666666'
    });

    s.addText(c.title || '', {
      x: x + 0.35, y: topY + 0.85, w: cardW - 0.7, h: 0.7,
      ...headlineStyle(theme, 22),
      color: theme.primary
    });

    s.addText(c.body || '', {
      x: x + 0.35, y: topY + 1.55, w: cardW - 0.7, h: cardH - 2.05,
      ...bodyStyle(theme, 14),
      color: '333333',
      valign: 'top',
      lineSpacingMultiple: 1.12
    });
  });

  if (imageFile) {
    // Decorative corner image stamp
    const bg = await coverImage(imageFile, 3.0, 1.7, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: SLIDE_W - 3.3, y: 0.2, w: 3.2, h: 1.7, transparency: 15 });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderImageCaption(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: theme.primary, transparency: L > 0.55 ? 35 : 60 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.background = { color: theme.primary };
  }

  // Caption pill
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.9, y: 5.9, w: SLIDE_W - 1.8, h: 1.3,
    fill: { color: 'FFFFFF', transparency: 12 },
    line: { color: 'FFFFFF', transparency: 100 }
  });

  s.addText(slide.title, {
    x: 1.2, y: 6.02, w: SLIDE_W - 2.4, h: 0.55,
    ...headlineStyle(theme, 34),
    color: 'FFFFFF'
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 1.2, y: 6.55, w: SLIDE_W - 2.4, h: 0.5,
      ...bodyStyle(theme, 14),
      color: 'FFFFFF'
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderTimeline(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 44),
    color: theme.primary
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.45, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const items = Array.isArray(slide.timeline_items) && slide.timeline_items.length
    ? slide.timeline_items
    : (slide.bullets || []).slice(0, 6).map((b, i) => ({ date_or_phase: `T${i + 1}`, label: b, detail: '' }));

  const max = Math.min(6, items.length);
  const leftX = 1.1;
  const rightX = SLIDE_W - 1.1;
  const yLine = 4.3;
  const lineW = rightX - leftX;

  // main line
  s.addShape(pptx.ShapeType.rect, {
    x: leftX, y: yLine, w: lineW, h: 0.07,
    fill: { color: 'E7E7EA' },
    line: { color: 'E7E7EA' }
  });

  for (let i = 0; i < max; i++) {
    const t = items[i];
    const x = leftX + (lineW * (max === 1 ? 0 : i / (max - 1)));
    const up = i % 2 === 0;

    // node
    s.addShape(pptx.ShapeType.ellipse, {
      x: x - 0.14, y: yLine - 0.16, w: 0.28, h: 0.28,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });
    s.addShape(pptx.ShapeType.ellipse, {
      x: x - 0.08, y: yLine - 0.10, w: 0.16, h: 0.16,
      fill: { color: 'FFFFFF' },
      line: { color: 'FFFFFF' }
    });

    const boxY = up ? 2.2 : 4.65;
    const stemY = up ? boxY + 1.15 : yLine + 0.07;
    const stemH = up ? (yLine - stemY) : (boxY - yLine - 0.35);

    // stem
    s.addShape(pptx.ShapeType.rect, {
      x: x - 0.01, y: stemY, w: 0.02, h: Math.max(0.2, stemH),
      fill: { color: 'DADAE0' },
      line: { color: 'DADAE0' }
    });

    // label box
    s.addShape(pptx.ShapeType.roundRect, {
      x: x - 1.15, y: boxY, w: 2.3, h: 1.05,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });
    s.addText(t.date_or_phase || '', {
      x: x - 1.05, y: boxY + 0.12, w: 2.1, h: 0.25,
      ...bodyStyle(theme, 11),
      color: '666666',
      align: 'center'
    });
    s.addText(t.label || '', {
      x: x - 1.05, y: boxY + 0.35, w: 2.1, h: 0.65,
      ...bodyStyle(theme, 13),
      color: theme.primary,
      align: 'center',
      valign: 'top'
    });
  }

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, 1.2, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 6.35, w: SLIDE_W, h: 1.15, transparency: 35 });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderKpiDashboard(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.6, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 42),
    color: theme.primary
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.35, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const kpis = Array.isArray(slide.kpis) && slide.kpis.length
    ? slide.kpis
    : (slide.bullets || []).slice(0, 6).map((b, i) => ({ label: `KPI ${i + 1}`, value: b, delta: '' }));

  const max = Math.min(6, kpis.length);
  const cols = 3;
  const rows = Math.ceil(max / cols);
  const gap = 0.35;
  const tileW = (SLIDE_W - 1.8 - gap * (cols - 1)) / cols;
  const tileH = 1.45;
  const startY = 2.15;

  for (let i = 0; i < max; i++) {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const x = 0.9 + c * (tileW + gap);
    const y = startY + r * (tileH + gap);
    const k = kpis[i];

    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: tileW, h: tileH,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });

    // accent
    s.addShape(pptx.ShapeType.rect, {
      x, y, w: 0.12, h: tileH,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });

    s.addText(k.value || '', {
      x: x + 0.25, y: y + 0.22, w: tileW - 0.5, h: 0.65,
      ...headlineStyle(theme, 30),
      color: theme.primary
    });
    s.addText(k.label || '', {
      x: x + 0.25, y: y + 0.92, w: tileW - 0.5, h: 0.35,
      ...bodyStyle(theme, 12),
      color: '666666'
    });
    if (k.delta) {
      s.addText(k.delta, {
        x: x + tileW - 1.1, y: y + 0.2, w: 0.9, h: 0.35,
        ...bodyStyle(theme, 11),
        color: theme.secondary,
        align: 'right'
      });
    }
  }

  // Footer bullets / notes strip
  const notes = (slide.bullets || []).slice(max, max + 4).filter(Boolean);
  if (notes.length) {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.9, y: 6.35, w: SLIDE_W - 1.8, h: 0.95,
      fill: { color: theme.primary, transparency: 92 },
      line: { color: 'E7E7EA' }
    });
    s.addText(notes.map(n => `• ${n}`).join('   '), {
      x: 1.1, y: 6.52, w: SLIDE_W - 2.2, h: 0.6,
      ...bodyStyle(theme, 12),
      color: theme.primary
    });
  }

  if (imageFile) {
    const bg = await coverImage(imageFile, 2.6, 2.6, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: SLIDE_W - 3.2, y: 0.15, w: 2.9, h: 2.9, transparency: 25 });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderTrafficLight(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 42),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.4, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const rows = Array.isArray(slide.status_items) && slide.status_items.length
    ? slide.status_items
    : (slide.bullets || []).slice(0, 6).map((b) => ({ item: b, status: 'yellow', owner: '', eta: '', blocker: '' }));

  const tableX = 0.9;
  const tableY = 2.2;
  const tableW = SLIDE_W - 1.8;
  const colW = [6.0, 2.0, 2.0, 1.3]; // item, owner, eta, status
  const rowH = 0.55;

  // Header row
  s.addShape(pptx.ShapeType.roundRect, {
    x: tableX, y: tableY, w: tableW, h: rowH,
    fill: { color: theme.primary, transparency: 0 },
    line: { color: theme.primary }
  });

  const headers = ['Workstream', 'Owner', 'ETA', 'Status'];
  let x = tableX;
  headers.forEach((h, i) => {
    s.addText(h, {
      x: x + 0.2, y: tableY + 0.12, w: colW[i] - 0.4, h: 0.3,
      ...bodyStyle(theme, 12),
      color: 'FFFFFF'
    });
    x += colW[i];
  });

  const colorFor = (st) => {
    const v = (st || '').toString().toLowerCase();
    if (v === 'green') return '2ECC71';
    if (v === 'red') return 'E74C3C';
    return 'F1C40F';
  };

  rows.slice(0, 10).forEach((r, i) => {
    const y = tableY + rowH + i * rowH;
    s.addShape(pptx.ShapeType.rect, {
      x: tableX, y, w: tableW, h: rowH,
      fill: { color: i % 2 === 0 ? 'F6F6F8' : 'FFFFFF' },
      line: { color: 'E7E7EA' }
    });

    const cells = [r.item || '', r.owner || '', r.eta || '', ''];
    let cx = tableX;
    for (let c = 0; c < 4; c++) {
      if (c === 3) {
        const dot = colorFor(r.status);
        s.addShape(pptx.ShapeType.ellipse, {
          x: cx + 0.55, y: y + 0.16, w: 0.23, h: 0.23,
          fill: { color: dot },
          line: { color: dot }
        });
      } else {
        s.addText(cells[c], {
          x: cx + 0.2, y: y + 0.12, w: colW[c] - 0.4, h: 0.3,
          ...bodyStyle(theme, 12),
          color: theme.primary
        });
      }
      cx += colW[c];
    }
  });

  // Blockers
  const blockers = rows.filter(r => (r.blocker || '').trim()).slice(0, 3);
  if (blockers.length) {
    s.addShape(pptx.ShapeType.roundRect, {
      x: 0.9, y: 6.45, w: SLIDE_W - 1.8, h: 0.8,
      fill: { color: 'FFF6E5' },
      line: { color: 'F3E1B0' }
    });
    s.addText(`Blockers: ${blockers.map(b => b.blocker).join(' • ')}`, {
      x: 1.1, y: 6.58, w: SLIDE_W - 2.2, h: 0.55,
      ...bodyStyle(theme, 11),
      color: theme.primary
    });
  }

  if (imageFile) {
    const bg = await coverImage(imageFile, 3.0, 1.6, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: SLIDE_W - 3.25, y: 0.15, w: 3.1, h: 1.6, transparency: 35 });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderTable(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 42),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.4, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  let headers = slide.table?.headers || null;
  let rows = slide.table?.rows || null;

  if (!headers || !rows) {
    // Attempt to derive a 2-col table from bullets: "Key: Value".
    const pairs = (slide.bullets || []).map(b => {
      const m = String(b).split(':');
      if (m.length >= 2) return [m[0].trim(), m.slice(1).join(':').trim()];
      return null;
    }).filter(Boolean);
    headers = ['Item', 'Detail'];
    rows = pairs.length ? pairs.slice(0, 10) : [['—', '—']];
  }

  const cols = headers.length;
  const x = 0.9;
  const y = 2.2;
  const w = SLIDE_W - 1.8;
  const h = 4.9;
  const rowH = Math.min(0.55, h / (Math.max(6, rows.length + 1)));
  const colW = w / cols;

  // Header background
  s.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h: rowH,
    fill: { color: theme.primary },
    line: { color: theme.primary }
  });

  for (let c = 0; c < cols; c++) {
    s.addText(headers[c], {
      x: x + c * colW + 0.15, y: y + 0.12, w: colW - 0.3, h: 0.3,
      ...bodyStyle(theme, 12),
      color: 'FFFFFF'
    });
  }

  rows.slice(0, 12).forEach((r, ri) => {
    const yy = y + rowH + ri * rowH;
    s.addShape(pptx.ShapeType.rect, {
      x, y: yy, w, h: rowH,
      fill: { color: ri % 2 === 0 ? 'F6F6F8' : 'FFFFFF' },
      line: { color: 'E7E7EA' }
    });
    for (let c = 0; c < cols; c++) {
      s.addText(String(r[c] ?? ''), {
        x: x + c * colW + 0.15, y: yy + 0.12, w: colW - 0.3, h: 0.3,
        ...bodyStyle(theme, 12),
        color: theme.primary
      });
    }
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderPricing(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 42),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.4, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const pricing = slide.pricing || { currency: '', plans: [], notes: '' };
  const plans = Array.isArray(pricing.plans) && pricing.plans.length
    ? pricing.plans
    : [
      { name: 'Basic', price: '—', period: 'mo', bullets: (slide.bullets || []).slice(0, 3), highlight: false },
      { name: 'Pro', price: '—', period: 'mo', bullets: (slide.bullets || []).slice(3, 6), highlight: true },
      { name: 'Enterprise', price: '—', period: 'yr', bullets: (slide.bullets || []).slice(6, 9), highlight: false },
    ];

  const cols = Math.min(3, plans.length);
  const gap = 0.35;
  const cardW = (SLIDE_W - 1.8 - gap * (cols - 1)) / cols;
  const cardH = 4.9;
  const y = 2.05;

  for (let i = 0; i < cols; i++) {
    const p = plans[i];
    const x = 0.9 + i * (cardW + gap);
    const isHot = Boolean(p.highlight);

    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: cardW, h: cardH,
      fill: { color: isHot ? 'F6FBFF' : 'F6F6F8' },
      line: { color: isHot ? theme.secondary : 'E7E7EA', width: isHot ? 2 : 1 }
    });

    if (isHot) {
      s.addShape(pptx.ShapeType.roundRect, {
        x: x + cardW - 1.35, y: y + 0.25, w: 1.1, h: 0.32,
        fill: { color: theme.secondary },
        line: { color: theme.secondary }
      });
      s.addText('RECOMMENDED', {
        x: x + cardW - 1.33, y: y + 0.31, w: 1.06, h: 0.22,
        ...bodyStyle(theme, 8),
        color: 'FFFFFF',
        align: 'center'
      });
    }

    s.addText(p.name, {
      x: x + 0.35, y: y + 0.55, w: cardW - 0.7, h: 0.5,
      ...headlineStyle(theme, 22),
      color: theme.primary
    });

    const priceLine = `${pricing.currency || ''} ${p.price}`.trim();
    s.addText(priceLine, {
      x: x + 0.35, y: y + 1.1, w: cardW - 0.7, h: 0.7,
      ...headlineStyle(theme, 36),
      color: theme.primary
    });
    s.addText(p.period ? `/${p.period}` : '', {
      x: x + 0.35, y: y + 1.72, w: cardW - 0.7, h: 0.3,
      ...bodyStyle(theme, 12),
      color: '666666'
    });

    const blt = (p.bullets || []).slice(0, 6).map(b => `• ${b}`).join('\n') || '• —';
    s.addText(blt, {
      x: x + 0.35, y: y + 2.15, w: cardW - 0.7, h: cardH - 2.45,
      ...bodyStyle(theme, 13),
      color: theme.primary,
      valign: 'top',
      lineSpacingMultiple: 1.12
    });
  }

  if (pricing.notes) {
    s.addText(pricing.notes, {
      x: 0.9, y: 6.95, w: SLIDE_W - 1.8, h: 0.35,
      ...bodyStyle(theme, 11),
      color: '666666'
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderComparisonMatrix(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 42),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.4, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  let m = slide.matrix;
  if (!m || !Array.isArray(m.x_labels) || !Array.isArray(m.y_labels) || !Array.isArray(m.cells)) {
    const xs = ['Option A', 'Option B', 'Option C'];
    const ys = (slide.bullets || []).slice(0, 5).map((b, i) => `Row ${i + 1}`) || ['Row 1'];
    const cells = ys.map(() => xs.map(() => '—'));
    m = { x_labels: xs, y_labels: ys, cells };
  }

  const xLabels = m.x_labels.slice(0, 5);
  const yLabels = m.y_labels.slice(0, 7);
  const cols = xLabels.length + 1; // include row header
  const rows = yLabels.length + 1; // include col header

  const x0 = 0.9;
  const y0 = 2.15;
  const w = SLIDE_W - 1.8;
  const h = 5.15;
  const colW = w / cols;
  const rowH = h / rows;

  // Header row background
  s.addShape(pptx.ShapeType.roundRect, {
    x: x0, y: y0, w, h: rowH,
    fill: { color: theme.primary },
    line: { color: theme.primary }
  });

  // Left header column background
  s.addShape(pptx.ShapeType.rect, {
    x: x0, y: y0, w: colW, h,
    fill: { color: 'F6F6F8' },
    line: { color: 'E7E7EA' }
  });

  // Column labels
  for (let c = 0; c < xLabels.length; c++) {
    s.addText(xLabels[c], {
      x: x0 + (c + 1) * colW + 0.1,
      y: y0 + 0.12,
      w: colW - 0.2,
      h: rowH - 0.2,
      ...bodyStyle(theme, 12),
      color: 'FFFFFF',
      align: 'center',
      valign: 'mid'
    });
  }

  for (let r = 0; r < yLabels.length; r++) {
    const yy = y0 + (r + 1) * rowH;

    // Row striping
    s.addShape(pptx.ShapeType.rect, {
      x: x0 + colW,
      y: yy,
      w: w - colW,
      h: rowH,
      fill: { color: r % 2 === 0 ? 'FFFFFF' : 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });

    // Row label
    s.addText(yLabels[r], {
      x: x0 + 0.1,
      y: yy + 0.12,
      w: colW - 0.2,
      h: rowH - 0.2,
      ...bodyStyle(theme, 12),
      color: theme.primary,
      valign: 'mid'
    });

    const rowCells = (m.cells[r] || []).slice(0, xLabels.length);
    for (let c = 0; c < xLabels.length; c++) {
      s.addText(String(rowCells[c] ?? ''), {
        x: x0 + (c + 1) * colW + 0.1,
        y: yy + 0.12,
        w: colW - 0.2,
        h: rowH - 0.2,
        ...bodyStyle(theme, 11),
        color: theme.primary,
        align: 'center',
        valign: 'mid'
      });
    }
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderProcessSteps(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title, {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 44),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.45, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const steps = Array.isArray(slide.steps) && slide.steps.length
    ? slide.steps
    : (slide.bullets || []).slice(0, 5).map((b, i) => ({ title: `Step ${i + 1}`, detail: b }));

  const max = Math.min(5, steps.length);
  const x0 = 1.0;
  const y0 = 2.55;
  const w = SLIDE_W - 2.0;
  const gap = 0.25;
  const boxW = (w - gap * (max - 1)) / max;
  const boxH = 3.4;

  for (let i = 0; i < max; i++) {
    const st = steps[i];
    const x = x0 + i * (boxW + gap);

    s.addShape(pptx.ShapeType.roundRect, {
      x, y: y0, w: boxW, h: boxH,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });

    // Number circle
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.25, y: y0 + 0.25, w: 0.45, h: 0.45,
      fill: { color: theme.secondary },
      line: { color: theme.secondary }
    });
    s.addText(String(i + 1), {
      x: x + 0.25, y: y0 + 0.33, w: 0.45, h: 0.3,
      ...bodyStyle(theme, 14),
      color: 'FFFFFF',
      align: 'center'
    });

    s.addText(st.title || '', {
      x: x + 0.25, y: y0 + 0.8, w: boxW - 0.5, h: 0.5,
      ...headlineStyle(theme, 16),
      color: theme.primary
    });

    s.addText(st.detail || '', {
      x: x + 0.25, y: y0 + 1.25, w: boxW - 0.5, h: boxH - 1.4,
      ...bodyStyle(theme, 12),
      color: '333333',
      valign: 'top',
      lineSpacingMultiple: 1.12
    });

    if (i < max - 1) {
      // Connector arrow
      s.addShape(pptx.ShapeType.rightArrow, {
        x: x + boxW - 0.02, y: y0 + 1.55, w: gap + 0.06, h: 0.35,
        fill: { color: theme.secondary, transparency: 25 },
        line: { color: theme.secondary, transparency: 100 }
      });
    }
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderTeamGrid(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title || 'Team', {
    x: 0.9, y: 0.65, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 44),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.45, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const people = Array.isArray(slide.people) && slide.people.length
    ? slide.people
    : (slide.bullets || []).slice(0, 6).map((b, i) => ({ name: `Person ${i + 1}`, role: '', bio: b }));

  const max = Math.min(6, people.length);
  const cols = 3;
  const rows = Math.ceil(max / cols);
  const gap = 0.35;
  const cardW = (SLIDE_W - 1.8 - gap * (cols - 1)) / cols;
  const cardH = 2.15;
  const y0 = 2.25;

  for (let i = 0; i < max; i++) {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const x = 0.9 + c * (cardW + gap);
    const y = y0 + r * (cardH + gap);
    const p = people[i];

    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: cardW, h: cardH,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });

    // Avatar placeholder
    s.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.25, y: y + 0.35, w: 0.65, h: 0.65,
      fill: { color: theme.secondary, transparency: 35 },
      line: { color: theme.secondary, transparency: 60 }
    });

    s.addText(p.name || '', {
      x: x + 1.0, y: y + 0.25, w: cardW - 1.2, h: 0.4,
      ...headlineStyle(theme, 16),
      color: theme.primary
    });

    s.addText(p.role || '', {
      x: x + 1.0, y: y + 0.65, w: cardW - 1.2, h: 0.35,
      ...bodyStyle(theme, 11),
      color: '666666'
    });

    s.addText(p.bio || '', {
      x: x + 0.25, y: y + 1.05, w: cardW - 0.5, h: cardH - 1.2,
      ...bodyStyle(theme, 11),
      color: '333333',
      valign: 'top',
      lineSpacingMultiple: 1.12
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderLogoWall(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title || 'Clients & Partners', {
    x: 0.9, y: 0.75, w: SLIDE_W - 1.8, h: 0.8,
    ...headlineStyle(theme, 44),
    color: theme.primary
  });
  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 1.55, w: SLIDE_W - 1.8, h: 0.5,
      ...bodyStyle(theme, 14),
      color: '666666'
    });
  }

  const items = (slide.logo_items || slide.bullets || []).filter(Boolean).slice(0, 24);
  const x0 = 0.9;
  const y0 = 2.35;
  const w = SLIDE_W - 1.8;
  const pillH = 0.5;
  const gapX = 0.2;
  const gapY = 0.2;

  let x = x0;
  let y = y0;
  items.forEach((name) => {
    const text = String(name);
    // Rough width heuristic: 0.12" per char + padding
    const pillW = Math.min(3.4, Math.max(1.3, 0.12 * text.length + 0.7));
    if (x + pillW > x0 + w) {
      x = x0;
      y += pillH + gapY;
    }
    if (y + pillH > SLIDE_H - 0.8) return;

    s.addShape(pptx.ShapeType.roundRect, {
      x, y, w: pillW, h: pillH,
      fill: { color: 'F6F6F8' },
      line: { color: 'E7E7EA' }
    });
    s.addText(text, {
      x: x + 0.2, y: y + 0.12, w: pillW - 0.4, h: 0.3,
      ...bodyStyle(theme, 12),
      color: theme.primary,
      align: 'center'
    });

    x += pillW + gapX;
  });

  // Accent corner
  s.addShape(pptx.ShapeType.roundRect, {
    x: SLIDE_W - 2.2, y: SLIDE_H - 0.9, w: 1.3, h: 0.45,
    fill: { color: theme.secondary },
    line: { color: theme.secondary }
  });
  s.addText('PROOF', {
    x: SLIDE_W - 2.2, y: SLIDE_H - 0.82, w: 1.3, h: 0.3,
    ...bodyStyle(theme, 11),
    color: 'FFFFFF',
    align: 'center'
  });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderCTA(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
      fill: { color: theme.primary, transparency: L > 0.55 ? 40 : 65 },
      line: { color: theme.primary, transparency: 100 }
    });
  } else {
    s.background = { color: theme.primary };
  }

  const headline = slide.cta?.headline || slide.title || 'Next steps';
  const primary = slide.cta?.primary_action || (slide.bullets?.[0] ?? 'Book a meeting');
  const secondary = slide.cta?.secondary_action || (slide.bullets?.[1] ?? 'Request a proposal');

  s.addText(headline, {
    x: 0.9, y: 2.0, w: SLIDE_W - 1.8, h: 1.6,
    ...headlineStyle(theme, 64),
    color: 'FFFFFF'
  });

  if (slide.subtitle) {
    s.addText(slide.subtitle, {
      x: 0.9, y: 3.55, w: SLIDE_W - 1.8, h: 0.7,
      ...bodyStyle(theme, 18),
      color: 'FFFFFF'
    });
  }

  // Primary button
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.9, y: 5.05, w: 4.6, h: 0.75,
    fill: { color: theme.secondary },
    line: { color: theme.secondary }
  });
  s.addText(primary, {
    x: 1.15, y: 5.2, w: 4.1, h: 0.45,
    ...headlineStyle(theme, 20),
    color: 'FFFFFF'
  });

  // Secondary
  s.addShape(pptx.ShapeType.roundRect, {
    x: 0.9, y: 5.95, w: 4.6, h: 0.65,
    fill: { color: 'FFFFFF', transparency: 15 },
    line: { color: 'FFFFFF', transparency: 60 }
  });
  s.addText(secondary, {
    x: 1.15, y: 6.08, w: 4.1, h: 0.4,
    ...bodyStyle(theme, 16),
    color: 'FFFFFF'
  });

  // Footer contact placeholder
  const contact = slide.bullets?.slice(2, 5).filter(Boolean).join(' • ');
  if (contact) {
    s.addText(contact, {
      x: 0.9, y: 6.85, w: SLIDE_W - 1.8, h: 0.4,
      ...bodyStyle(theme, 12),
      color: 'FFFFFF'
    });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}


// ---------- Additional Business Layouts ----------
function safeArr(a) { return Array.isArray(a) ? a : []; }

function addTitleBlock(slide, theme, title, subtitle, color='111111') {
  const t = title || '';
  slide.addText(t, { x: 0.9, y: 0.55, w: SLIDE_W - 1.8, h: 0.8, ...headlineStyle(theme, theme.style?.h2Size ?? 34), color });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.9, y: 1.35, w: SLIDE_W - 1.8, h: 0.5, ...bodyStyle(theme, theme.style?.bodySize ?? 16), color });
  }
}

async function renderSWOT(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  // Optional soft background
  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x: 0, y: 0, w: SLIDE_W, h: SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:SLIDE_H, fill:{ color: theme.primary, transparency: overlayTransparency(theme, L) }, line:{ color: theme.primary, transparency: 100 } });
  } else {
    s.background = { color: theme.primary };
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:SLIDE_H, fill:{ color:'FFFFFF', transparency: 6 }, line:{ color:'FFFFFF', transparency: 100 }});
  }

  addTitleBlock(s, theme, slide.title, slide.subtitle, 'FFFFFF');

  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme, { fillTransparency: theme.style?.panelTransparency ?? 12 });

  const gridX = 0.9, gridY = 2.05, gridW = SLIDE_W - 1.8, gridH = 5.05;
  const colW = gridW/2, rowH = gridH/2;

  const sw = slide.swot || {};
  const cells = [
    { label: 'Strengths', items: safeArr(sw.strengths), x: gridX, y: gridY },
    { label: 'Weaknesses', items: safeArr(sw.weaknesses), x: gridX + colW, y: gridY },
    { label: 'Opportunities', items: safeArr(sw.opportunities), x: gridX, y: gridY + rowH },
    { label: 'Threats', items: safeArr(sw.threats), x: gridX + colW, y: gridY + rowH }
  ];

  for (const c of cells) {
    s.addShape(panelType, { x: c.x, y: c.y, w: colW - 0.15, h: rowH - 0.15, ...pStyle });
    s.addText(c.label, { x: c.x + 0.25, y: c.y + 0.2, w: colW - 0.65, h: 0.4, ...headlineStyle(theme, 18), color: '111111' });
    const lines = c.items.slice(0, 6).map(v => `• ${v}`);
    s.addText(lines.join('\n') || '• (add items)', { x: c.x + 0.25, y: c.y + 0.7, w: colW - 0.65, h: rowH - 1.0, ...bodyStyle(theme, 14), color: '111111' });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderFunnel(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x:0, y:0, w:SLIDE_W, h:SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:SLIDE_H, fill:{ color: theme.primary, transparency: overlayTransparency(theme, L) }, line:{ color: theme.primary, transparency: 100 } });
  } else {
    s.background = { color: theme.primary };
  }

  addTitleBlock(s, theme, slide.title, slide.subtitle, 'FFFFFF');

  const stages = safeArr(slide.funnel).slice(0, 6);
  const x = 2.2, y = 2.1, w = SLIDE_W - 4.4, h = 4.9;
  const stepH = h / Math.max(1, stages.length);

  for (let i=0;i<stages.length;i++) {
    const st = stages[i] || {};
    const inset = (i * 0.35);
    s.addShape(pptx.ShapeType.roundRect, {
      x: x + inset, y: y + i*stepH,
      w: w - inset*2, h: stepH - 0.18,
      fill: { color: 'FFFFFF', transparency: 10 },
      line: { color: theme.secondary, transparency: 25, width: 1 }
    });
    s.addText(st.label || `Stage ${i+1}`, { x: x + inset + 0.35, y: y + i*stepH + 0.18, w: w - inset*2 - 0.7, h: 0.35, ...headlineStyle(theme, 16), color:'111111' });
    s.addText(st.value || '', { x: x + inset + 0.35, y: y + i*stepH + 0.62, w: w - inset*2 - 0.7, h: 0.4, ...bodyStyle(theme, 14), color:'111111' });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderNowNextLater(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x:0, y:0, w:SLIDE_W, h:SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:SLIDE_H, fill:{ color: theme.primary, transparency: overlayTransparency(theme, L) }, line:{ color: theme.primary, transparency: 100 } });
  } else {
    s.background = { color: theme.primary };
  }

  addTitleBlock(s, theme, slide.title, slide.subtitle, 'FFFFFF');

  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme, { fillTransparency: 10 });

  const data = slide.now_next_later || { now: [], next: [], later: [] };
  const cols = [
    { label: 'Now', items: safeArr(data.now), x: 0.9 },
    { label: 'Next', items: safeArr(data.next), x: 4.78 },
    { label: 'Later', items: safeArr(data.later), x: 8.66 }
  ];

  for (const c of cols) {
    s.addShape(panelType, { x: c.x, y: 2.05, w: 3.77, h: 5.1, ...pStyle });
    s.addShape(pptx.ShapeType.rect, { x: c.x, y: 2.05, w: 3.77, h: 0.12, fill:{ color: theme.secondary }, line:{ color: theme.secondary }});
    s.addText(c.label, { x: c.x + 0.25, y: 2.25, w: 3.27, h: 0.4, ...headlineStyle(theme, 18), color:'111111' });
    const lines = c.items.slice(0, 8).map(v => `• ${v}`);
    s.addText(lines.join('\n') || '• (add items)', { x: c.x + 0.25, y: 2.7, w: 3.27, h: 4.2, ...bodyStyle(theme, 14), color:'111111' });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderOKR(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  // Clean white canvas with optional accent image strip
  s.background = { color: 'FFFFFF' };
  if (imageFile) {
    const strip = await coverImage(imageFile, SLIDE_W, 2.0, tmpDir, imgCropCache);
    s.addImage({ path: strip, x: 0, y: 0, w: SLIDE_W, h: 2.0 });
    const L = await getLuminance(strip);
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:2.0, fill:{ color: theme.primary, transparency: overlayTransparency(theme, L) }, line:{ color: theme.primary, transparency:100 }});
    addTitleBlock(s, theme, slide.title, slide.subtitle, 'FFFFFF');
  } else {
    addTitleBlock(s, theme, slide.title, slide.subtitle, '111111');
  }

  const okrs = safeArr(slide.okrs).slice(0, 4);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme, { fillTransparency: 0, lineTransparency: 60 });

  let y = 2.2;
  for (let i=0;i<okrs.length;i++) {
    const o = okrs[i] || {};
    const h = 1.25;
    s.addShape(panelType, { x: 0.9, y, w: SLIDE_W - 1.8, h, ...pStyle });
    s.addShape(pptx.ShapeType.rect, { x: 0.9, y, w: 0.12, h, fill:{ color: theme.secondary }, line:{ color: theme.secondary }});
    s.addText(o.objective || `Objective ${i+1}`, { x: 1.1, y: y + 0.12, w: SLIDE_W - 2.2, h: 0.35, ...headlineStyle(theme, 18), color:'111111' });
    const kr = safeArr(o.key_results).slice(0, 5).map(v => `• ${v}`);
    s.addText(kr.join('\n') || '• (add key results)', { x: 1.1, y: y + 0.48, w: SLIDE_W - 2.2, h: 0.75, ...bodyStyle(theme, 14), color:'111111' });
    y += h + 0.25;
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderCaseStudy(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x:0, y:0, w:SLIDE_W, h:SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:SLIDE_H, fill:{ color: theme.primary, transparency: overlayTransparency(theme, L) }, line:{ color: theme.primary, transparency: 100 } });
    addTitleBlock(s, theme, slide.title, slide.subtitle, 'FFFFFF');
  } else {
    s.background = { color: 'FFFFFF' };
    addTitleBlock(s, theme, slide.title, slide.subtitle, '111111');
  }

  const cs = slide.case_study || {};
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme, { fillTransparency: 6 });

  // left: context, right: results
  s.addShape(panelType, { x: 0.9, y: 2.05, w: 6.15, h: 5.1, ...pStyle });
  s.addShape(panelType, { x: 7.1, y: 2.05, w: 5.33, h: 5.1, ...pStyle });

  s.addText(`Client`, { x: 1.15, y: 2.25, w: 5.8, h: 0.35, ...headlineStyle(theme, 16), color: '111111' });
  s.addText(cs.client || '(client name)', { x: 1.15, y: 2.62, w: 5.8, h: 0.4, ...bodyStyle(theme, 14), color:'111111' });

  s.addText(`Challenge`, { x: 1.15, y: 3.15, w: 5.8, h: 0.35, ...headlineStyle(theme, 16), color: '111111' });
  s.addText(cs.challenge || '(challenge)', { x: 1.15, y: 3.52, w: 5.8, h: 0.9, ...bodyStyle(theme, 14), color:'111111' });

  s.addText(`Approach`, { x: 1.15, y: 4.55, w: 5.8, h: 0.35, ...headlineStyle(theme, 16), color: '111111' });
  const ap = safeArr(cs.approach).slice(0, 6).map(v => `• ${v}`).join('\n') || '• (add approach steps)';
  s.addText(ap, { x: 1.15, y: 4.92, w: 5.8, h: 2.0, ...bodyStyle(theme, 14), color:'111111' });

  s.addText(`Results`, { x: 7.35, y: 2.25, w: 4.9, h: 0.35, ...headlineStyle(theme, 16), color: '111111' });
  const res = safeArr(cs.results).slice(0, 8).map(v => `• ${v}`).join('\n') || '• (add results)';
  s.addText(res, { x: 7.35, y: 2.62, w: 4.9, h: 4.3, ...bodyStyle(theme, 14), color:'111111' });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

function minMax(vals) {
  const v = (Array.isArray(vals) ? vals : []).map(Number).filter(Number.isFinite);
  if (!v.length) return { min: 0, max: 1 };
  return { min: Math.min(...v), max: Math.max(...v) };
}

async function renderChartBar(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  addTitleBlock(s, theme, slide.title, slide.subtitle, '111111');

  const chart = slide.chart || { labels: [], values: [], value_suffix: '' };
  const labels = safeArr(chart.labels).slice(0, 8);
  const values = safeArr(chart.values).slice(0, labels.length).map(Number);
  const suf = chart.value_suffix || '';

  const area = { x: 1.1, y: 2.2, w: SLIDE_W - 2.2, h: 4.9 };
  // axis
  s.addShape(pptx.ShapeType.line, { x: area.x, y: area.y + area.h, w: area.w, h: 0, line: { color: '999999', width: 1 } });

  const { min, max } = minMax(values);
  const span = (max - min) || 1;
  const barW = area.w / Math.max(1, values.length) - 0.2;

  for (let i=0;i<values.length;i++) {
    const v = Number.isFinite(values[i]) ? values[i] : 0;
    const norm = (v - min) / span;
    const bh = Math.max(0.2, norm * (area.h - 0.6));
    const bx = area.x + i*(barW+0.2);
    const by = area.y + area.h - bh;
    s.addShape(pptx.ShapeType.rect, { x: bx, y: by, w: barW, h: bh, fill: { color: theme.secondary, transparency: 15 }, line: { color: theme.secondary, transparency: 35 }});
    s.addText(`${v}${suf}`, { x: bx, y: by - 0.28, w: barW, h: 0.25, fontFace: theme.bodyFont, fontSize: 11, color: '111111', align: 'center' });
    s.addText(labels[i] || '', { x: bx, y: area.y + area.h + 0.05, w: barW, h: 0.4, fontFace: theme.bodyFont, fontSize: 11, color: '555555', align: 'center' });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderChartLine(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  addTitleBlock(s, theme, slide.title, slide.subtitle, '111111');

  const chart = slide.chart || { labels: [], values: [], value_suffix: '' };
  const labels = safeArr(chart.labels).slice(0, 10);
  const values = safeArr(chart.values).slice(0, labels.length).map(Number);
  const suf = chart.value_suffix || '';

  const area = { x: 1.1, y: 2.2, w: SLIDE_W - 2.2, h: 4.9 };

  const { min, max } = minMax(values);
  const span = (max - min) || 1;

  const pts = values.map((v, i) => {
    const norm = (Number.isFinite(v) ? v : min) - min;
    const y = area.y + area.h - (norm / span) * (area.h - 0.6) - 0.3;
    const x = area.x + (i/(Math.max(1, values.length-1))) * area.w;
    return { x, y, v };
  });

  // grid line
  s.addShape(pptx.ShapeType.line, { x: area.x, y: area.y + area.h, w: area.w, h: 0, line: { color: '999999', width: 1 } });

  // draw segments
  for (let i=0;i<pts.length-1;i++) {
    const a = pts[i], b = pts[i+1];
    s.addShape(pptx.ShapeType.line, { x: a.x, y: a.y, w: (b.x-a.x), h: (b.y-a.y), line: { color: theme.secondary, width: 2 }});
  }

  // points + labels
  for (let i=0;i<pts.length;i++) {
    const p = pts[i];
    s.addShape(pptx.ShapeType.ellipse, { x: p.x-0.07, y: p.y-0.07, w: 0.14, h: 0.14, fill: { color: theme.secondary }, line: { color: theme.secondary }});
    s.addText(`${p.v}${suf}`, { x: p.x-0.35, y: p.y-0.35, w: 0.7, h: 0.25, fontFace: theme.bodyFont, fontSize: 10, color: '111111', align: 'center' });
    s.addText(labels[i] || '', { x: p.x-0.6, y: area.y + area.h + 0.05, w: 1.2, h: 0.4, fontFace: theme.bodyFont, fontSize: 11, color: '555555', align: 'center' });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderOrgChart(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };
  addTitleBlock(s, theme, slide.title, slide.subtitle, '111111');

  const oc = slide.org_chart || { head: '', reports: [] };
  const head = oc.head || '(Head)';
  const reports = safeArr(oc.reports).slice(0, 8);

  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme, { fillTransparency: 0, lineTransparency: 55 });

  // head box
  s.addShape(panelType, { x: 4.4, y: 2.1, w: 4.5, h: 0.9, ...pStyle });
  s.addShape(pptx.ShapeType.rect, { x: 4.4, y: 2.1, w: 4.5, h: 0.1, fill:{ color: theme.secondary }, line:{ color: theme.secondary }});
  s.addText(head, { x: 4.55, y: 2.35, w: 4.2, h: 0.5, ...headlineStyle(theme, 18), color:'111111', align:'center' });

  // connector line
  s.addShape(pptx.ShapeType.line, { x: 6.65, y: 3.0, w: 0, h: 0.6, line: { color:'888888', width: 1 }});

  // report boxes
  const rowY = 3.8;
  const boxW = (SLIDE_W - 2.0) / Math.max(2, Math.min(4, reports.length));
  const cols = Math.min(4, Math.max(2, reports.length));
  const startX = (SLIDE_W - (cols*boxW)) / 2;

  for (let i=0;i<reports.length;i++) {
    const col = i % cols;
    const r = Math.floor(i / cols);
    const x = startX + col*boxW + 0.1;
    const y = rowY + r*1.2;
    s.addShape(panelType, { x, y, w: boxW - 0.2, h: 0.85, ...pStyle });
    s.addText(reports[i], { x: x + 0.15, y: y + 0.25, w: boxW - 0.5, h: 0.45, ...bodyStyle(theme, 14), color:'111111', align:'center' });
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderFAQ(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);

  if (imageFile) {
    const bg = await coverImage(imageFile, SLIDE_W, SLIDE_H, tmpDir, imgCropCache);
    s.addImage({ path: bg, x:0, y:0, w:SLIDE_W, h:SLIDE_H });
    const L = await getLuminance(bg);
    s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:SLIDE_W, h:SLIDE_H, fill:{ color: theme.primary, transparency: overlayTransparency(theme, L) }, line:{ color: theme.primary, transparency: 100 } });
    addTitleBlock(s, theme, slide.title, slide.subtitle, 'FFFFFF');
  } else {
    s.background = { color: 'FFFFFF' };
    addTitleBlock(s, theme, slide.title, slide.subtitle, '111111');
  }

  const items = safeArr(slide.faq).slice(0, 6);
  const panelType = panelShapeType(pptx, theme);
  const pStyle = panelStyle(theme, { fillTransparency: 6 });

  let y = 2.1;
  for (let i=0;i<items.length;i++) {
    const it = items[i] || {};
    const h = 0.9;
    s.addShape(panelType, { x: 0.9, y, w: SLIDE_W - 1.8, h, ...pStyle });
    s.addText(`Q: ${it.q || '(question)'}`, { x: 1.15, y: y + 0.12, w: SLIDE_W - 2.3, h: 0.3, ...headlineStyle(theme, 14), color:'111111' });
    s.addText(`A: ${it.a || '(answer)'}`, { x: 1.15, y: y + 0.42, w: SLIDE_W - 2.3, h: 0.4, ...bodyStyle(theme, 13), color:'111111' });
    y += h + 0.18;
  }

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}

async function renderAppendix(pptx, slide, theme, imageFile, tmpDir, imgCropCache) {
  // Appendix is a clean text-heavy slide (good for extra notes)
  const s = pptx.addSlide();
  addSlideFrame(s, pptx, theme);
  s.background = { color: 'FFFFFF' };

  s.addText(slide.title || 'Appendix', { x: 0.9, y: 0.6, w: SLIDE_W - 1.8, h: 0.7, ...headlineStyle(theme, 34), color:'111111' });
  if (slide.subtitle) {
    s.addText(slide.subtitle, { x: 0.9, y: 1.3, w: SLIDE_W - 1.8, h: 0.4, ...bodyStyle(theme, 14), color:'555555' });
  }
  const bullets = safeArr(slide.bullets).slice(0, 12).map(v => `• ${v}`);
  s.addText(bullets.join('\n') || '• (add appendix bullets)', { x: 0.9, y: 2.1, w: SLIDE_W - 1.8, h: 5.2, ...bodyStyle(theme, 16), color:'111111' });

  if (slide.speaker_notes) s.addNotes(slide.speaker_notes);
}
