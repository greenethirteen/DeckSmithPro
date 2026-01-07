import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import multer from 'multer';
import fs from 'fs-extra';
import path from 'path';
import { fileURLToPath } from 'url';
import { nanoid } from 'nanoid';

import { extractBriefText } from './lib/brief.js';
import { planDeck, normalizePlan } from './lib/planner.js';
import { exportPptx } from './lib/pptx.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT || 8787;
const CLIENT_ORIGIN = process.env.CLIENT_ORIGIN || 'http://localhost:5173';

const app = express();

app.use(cors({
  origin: CLIENT_ORIGIN,
  credentials: false,
}));
app.use(express.json({ limit: '25mb' }));

const TMP_DIR = process.env.TMP_DIR || path.join(__dirname, '.tmp');
await fs.ensureDir(TMP_DIR);

// ---------------- Export jobs + SSE progress (thumbnails during export) ----------------
const exportJobs = new Map(); // id -> { id, createdAt, status, events: [], subscribers: Set, filePath, filename, error }
const JOB_TTL_MS = 1000 * 60 * 30; // 30 minutes

function sseInit(res) {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();
}

function sseSend(res, event, data) {
  res.write(`event: ${event}\n`);
  res.write(`data: ${JSON.stringify(data)}\n\n`);
}

function jobPush(job, event, payload) {
  const msg = { ts: Date.now(), event, payload };
  job.events.push(msg);
  if (job.events.length > 2000) job.events.shift();
  for (const sub of job.subscribers) {
    try { sseSend(sub, event, payload); } catch {}
  }
}

function cleanupJobs() {
  const now = Date.now();
  for (const [id, job] of exportJobs.entries()) {
    if (now - job.createdAt > JOB_TTL_MS) {
      if (job.filePath) fs.remove(job.filePath).catch(()=>{});
      exportJobs.delete(id);
    }
  }
}
setInterval(cleanupJobs, 60 * 1000).unref?.();


// Multer for brief file upload (PDF/DOC/DOCX/PPTX)
const upload = multer({
  storage: multer.diskStorage({
    destination: async (req, file, cb) => {
      await fs.ensureDir(TMP_DIR);
      cb(null, TMP_DIR);
    },
    filename: (req, file, cb) => {
      const safe = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
      cb(null, `${Date.now()}_${nanoid(6)}_${safe}`);
    }
  }),
  limits: { fileSize: 25 * 1024 * 1024 }, // 25MB
});

app.get('/api/health', (req, res) => {
  res.json({ ok: true, ts: new Date().toISOString() });
});

/**
 * POST /api/plan
 * multipart/form-data: { file?: PDF, text?: string, options?: JSON string }
 * Returns: { extractedText, plan }
 */
app.post('/api/plan', upload.single('file'), async (req, res) => {
  try {
    const extraText = (req.body.text || '').toString();
    const options = req.body.options ? JSON.parse(req.body.options) : {};
    let extractedText = '';

    if (req.file) {
      try {
        extractedText = await extractBriefText(req.file.path, req.file.originalname);
      } catch (e) {
        return res.status(400).json({ error: e?.message || 'Unsupported brief file type.' });
      }
    }

    const briefText = [extractedText, extraText].filter(Boolean).join('\n\n---\n\n').trim();
    if (!briefText) return res.status(400).json({ error: 'Provide a PDF and/or text.' });

    const rawPlan = await planDeck(briefText, options);
    const plan = normalizePlan(rawPlan, options);

    res.json({ extractedText, plan });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err?.message || 'Failed to create plan' });
  } finally {
    // Best-effort cleanup
    if (req.file?.path) fs.remove(req.file.path).catch(()=>{});
  }
});

/**
 * POST /api/export
 * JSON body: { plan, options }
 * Returns: PPTX binary
 */
app.post('/api/export', async (req, res) => {
  try {
    const { plan, options } = req.body || {};
    if (!plan) return res.status(400).json({ error: 'Missing plan' });

    const out = await exportPptx(plan, options || {}, { tmpDir: TMP_DIR });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${(plan.deck_title || 'Deck').replace(/[^a-zA-Z0-9._-]/g,'_')}.pptx"`);
    res.send(out);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err?.message || 'Failed to export pptx' });
  }
});


/**
 * POST /api/export_job
 * JSON body: { plan, options }
 * Returns: { jobId, totalSlides, filenameHint }
 */
app.post('/api/export_job', async (req, res) => {
  try {
    const { plan, options } = req.body || {};
    if (!plan) return res.status(400).json({ error: 'Missing plan' });

    const jobId = nanoid(12);
    const filenameHint = `${(plan.deck_title || 'Deck').replace(/[^a-zA-Z0-9._-]/g,'_')}.pptx`;
    const job = {
      id: jobId,
      createdAt: Date.now(),
      status: 'queued',
      events: [],
      subscribers: new Set(),
      filePath: null,
      filename: filenameHint,
      thumbDir: path.join(TMP_DIR, `thumbs_${jobId}`),
      error: null
    };
    exportJobs.set(jobId, job);

    // Kick off async export
    (async () => {
      try {
        job.status = 'running';
        await fs.ensureDir(job.thumbDir);
        // Clean old thumbs if any
        await fs.emptyDir(job.thumbDir).catch(()=>{});
        jobPush(job, 'meta', { jobId, totalSlides: Array.isArray(plan.slides) ? plan.slides.length : 0, filename: filenameHint });
        jobPush(job, 'status', { phase: 'starting', message: 'Starting export…' });

        const buf = await exportPptx(plan, options || {}, {
          tmpDir: TMP_DIR,
          onStatus: (st) => jobPush(job, 'status', st),
          onSlide: (sl) => jobPush(job, 'slide', sl),
          sofficeThumbs: true,
          sofficePath: process.env.SOFFICE_PATH || 'soffice',
          jobId,
          thumbDir: job.thumbDir,
          onThumbnail: ({ index }) => {
            jobPush(job, 'thumbnail', { index, url: `/api/export_job/${jobId}/thumb/${index}.png` });
          }
        });

        const outPath = path.join(TMP_DIR, `export_${jobId}.pptx`);
        await fs.writeFile(outPath, buf);
        job.filePath = outPath;
        job.status = 'done';
        jobPush(job, 'done', { jobId, filename: filenameHint });
      } catch (e) {
        job.status = 'error';
        job.error = e?.message || String(e);
        jobPush(job, 'error', { message: job.error });
      } finally {
        // End streams
        for (const sub of job.subscribers) {
          try { sub.end(); } catch {}
        }
        job.subscribers.clear();
      }
    })();

    res.json({ jobId, totalSlides: Array.isArray(plan.slides) ? plan.slides.length : 0, filenameHint });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err?.message || 'Failed to start export' });
  }
});

/**
 * GET /api/export_job/:id/stream
 * SSE stream of export progress events: meta, status, slide, done, error
 */
app.get('/api/export_job/:id/stream', (req, res) => {
  const id = req.params.id;
  const job = exportJobs.get(id);
  if (!job) return res.status(404).end();
  sseInit(res);

  // Replay existing events so UI can catch up
  for (const e of job.events) {
    sseSend(res, e.event, e.payload);
  }

  job.subscribers.add(res);
  req.on('close', () => job.subscribers.delete(res));
});

/**
 * GET /api/export_job/:id/pptx
 * Download PPTX for a finished job
 */
app.get('/api/export_job/:id/pptx', async (req, res) => {
  const id = req.params.id;
  const job = exportJobs.get(id);
  if (!job) return res.status(404).json({ error: 'Not found' });
  if (job.status !== 'done' || !job.filePath) return res.status(409).json({ error: 'Not ready' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  res.setHeader('Content-Disposition', `attachment; filename="${job.filename || `export_${id}.pptx`}"`);
  fs.createReadStream(job.filePath).pipe(res);
});


/**
 * GET /api/export_job/:id/thumb/:idx.png
 * Serves per-slide PNG thumbnails generated during export (LibreOffice).
 */
app.get('/api/export_job/:id/thumb/:idx', async (req, res) => {
  const id = req.params.id;
  const job = exportJobs.get(id);
  const idx = Number.parseInt((req.params.idx || '').toString().replace(/\.png$/i,''), 10);
  if (!job) return res.status(404).end();
  if (!Number.isFinite(idx) || idx < 0) return res.status(400).end();

  const base = `partial_${id}`; // matches partial PPTX base name in exportPptx
  const cand = idx === 0 ? [`${base}.png`, `${base}_0.png`] : [`${base}_${idx}.png`, `${base}-${idx}.png`, `${base} (${idx}).png`];
  let filePath = null;
  for (const f of cand) {
    const p = path.join(job.thumbDir, f);
    if (await fs.pathExists(p)) { filePath = p; break; }
  }
  if (!filePath) {
    // fallback: find any png that matches index pattern
    const files = await fs.readdir(job.thumbDir).catch(()=>[]);
    const pngs = files.filter(f => f.toLowerCase().endsWith('.png') && f.startsWith(base));
    const match = pngs.find(f => f.toLowerCase() === `${base.toLowerCase()}.png` ? idx===0 : f.toLowerCase().includes(`_${idx}.png`));
    if (match) filePath = path.join(job.thumbDir, match);
  }

  if (!filePath) return res.status(404).end();
  res.setHeader('Content-Type', 'image/png');
  fs.createReadStream(filePath).pipe(res);
});

// Serve production client build if you want to deploy as a single app
const CLIENT_DIST = path.join(__dirname, '..', 'client', 'dist');
if (fs.existsSync(CLIENT_DIST)) {
  app.use(express.static(CLIENT_DIST));
  app.get('*', (req, res) => res.sendFile(path.join(CLIENT_DIST, 'index.html')));
}

app.listen(PORT, () => {
  console.log(`DeckSmith Pro server running on http://localhost:${PORT}`);
  console.log(`CORS allowed origin: ${CLIENT_ORIGIN}`);
});
