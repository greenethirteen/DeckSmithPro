import fs from 'fs-extra';
import os from 'os';
import path from 'path';
import sharp from 'sharp';
import { spawn } from 'child_process';
import * as vega from 'vega';
import * as vegaLite from 'vega-lite';
import { loader as vegaLoader } from 'vega-loader';

const ICON_CACHE = new Map();
const ICON_CACHE_MAX = 200;

function cacheGet(map, key) {
  return map.has(key) ? map.get(key) : null;
}

function cacheSet(map, key, value) {
  map.set(key, value);
  if (map.size > ICON_CACHE_MAX) {
    const oldest = map.keys().next().value;
    map.delete(oldest);
  }
}

function readEnvFlag(name, fallback) {
  const raw = (process.env[name] || '').toString().trim().toLowerCase();
  if (raw === 'true') return true;
  if (raw === 'false') return false;
  return fallback;
}

function resolveFormat(format) {
  const forcePng = readEnvFlag('PPTX_GRAPHICS_FORCE_PNG', false);
  if (forcePng) return 'png';
  if (format) return format;
  const preferSvg = readEnvFlag('PPTX_GRAPHICS_PREFER_SVG', true);
  return preferSvg ? 'svg' : 'png';
}

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

export async function renderChart({ spec, format, width = 1200, height = 800 }) {
  const fmt = resolveFormat(format);
  const vlSpec = { ...(spec || {}) };
  if (!('width' in vlSpec)) vlSpec.width = width;
  if (!('height' in vlSpec)) vlSpec.height = height;

  const compiled = vegaLite.compile(vlSpec);
  const vgSpec = compiled?.spec || compiled;
  const view = new vega.View(vega.parse(vgSpec), {
    renderer: 'none',
    loader: vegaLoader(),
    logger: vega.logger(vega.Warn)
  }).initialize();

  const svg = await view.toSVG();
  if (fmt === 'svg') return { svg };
  return { png: await svgToPng({ svg, width, height }) };
}

export async function renderDiagram({ code, format, width = 1200, height = 800, theme = 'default' }) {
  const fmt = resolveFormat(format);
  const outExt = fmt === 'png' ? 'png' : 'svg';
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'decksmith-mermaid-'));
  const inputPath = path.join(tmpDir, 'diagram.mmd');
  const outputPath = path.join(tmpDir, `diagram.${outExt}`);
  const cliPath = process.env.MERMAID_CLI_PATH
    || path.join(process.cwd(), 'node_modules', '.bin', 'mmdc');

  try {
    await fs.writeFile(inputPath, code || '');
    if (!await fs.pathExists(cliPath)) {
      throw new Error('Mermaid CLI not available.');
    }
    const args = [
      '-i', inputPath,
      '-o', outputPath,
      '-t', theme,
      '-w', String(width),
      '-H', String(height)
    ];
    await spawnCmd(cliPath, args);
    if (fmt === 'png') {
      const png = await fs.readFile(outputPath);
      return { png };
    }
    const svg = await fs.readFile(outputPath, 'utf8');
    return { svg };
  } catch (err) {
    // Fallback: return a simple SVG with the source code as text.
    const svg = fallbackMermaidSvg(code || '', width, height);
    return fmt === 'png' ? { png: await svgToPng({ svg, width, height }) } : { svg };
  } finally {
    await fs.remove(tmpDir).catch(() => {});
  }
}

export async function getIcon({ name, format, width = 128, height = 128 }) {
  if (!name) throw new Error('Icon name is required.');
  const fmt = resolveFormat(format);
  const base = process.env.ICONIFY_API_BASE || 'https://api.iconify.design';
  const key = `${base}|${name}|${width}|${height}`;
  const cached = cacheGet(ICON_CACHE, key);
  if (cached) {
    return fmt === 'png' ? { png: await svgToPng({ svg: cached, width, height }) } : { svg: cached };
  }
  const url = `${base.replace(/\/$/, '')}/${name}.svg?width=${width}&height=${height}`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Iconify fetch failed (${res.status}) for ${name}`);
  }
  const svg = await res.text();
  cacheSet(ICON_CACHE, key, svg);
  return fmt === 'png' ? { png: await svgToPng({ svg, width, height }) } : { svg };
}

export async function svgToPng({ svg, width, height }) {
  if (!svg) throw new Error('SVG is required.');
  try {
    return await sharp(Buffer.from(svg)).resize(width, height).png().toBuffer();
  } catch (err) {
    const mod = await import('@resvg/resvg-js');
    const Resvg = mod.Resvg || mod.default?.Resvg;
    if (!Resvg) throw new Error('Resvg renderer not available.');
    const resvg = new Resvg(svg, { fitTo: { mode: 'width', value: width } });
    const png = resvg.render().asPng();
    return Buffer.from(png);
  }
}

export function toDataUri({ svg, png }) {
  if (svg) return `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
  if (png) return `data:image/png;base64,${Buffer.from(png).toString('base64')}`;
  throw new Error('svg or png is required for data URI.');
}

function fallbackMermaidSvg(code, width, height) {
  const safe = (code || '').replace(/[<>&]/g, (m) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[m]));
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">`,
    `<rect width="100%" height="100%" fill="#FFFFFF" stroke="#E5E5E5"/>`,
    `<text x="24" y="40" font-family="Arial" font-size="18" fill="#111111">Mermaid render unavailable</text>`,
    `<text x="24" y="70" font-family="Arial" font-size="12" fill="#666666">Install @mermaid-js/mermaid-cli or set MERMAID_CLI_PATH.</text>`,
    `<foreignObject x="24" y="90" width="${width - 48}" height="${height - 110}">`,
    `<div xmlns="http://www.w3.org/1999/xhtml" style="font-family: monospace; font-size: 11px; color: #444; white-space: pre-wrap;">${safe}</div>`,
    `</foreignObject>`,
    `</svg>`
  ].join('');
}
