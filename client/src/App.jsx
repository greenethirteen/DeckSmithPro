import React, { useEffect, useMemo, useState } from 'react'
import Dropzone from './components/Dropzone.jsx'
import PlanEditor from './components/PlanEditor.jsx'

async function apiPlan({ file, text, options }) {
  const fd = new FormData()
  if (file) fd.append('file', file)
  if (text) fd.append('text', text)
  fd.append('options', JSON.stringify(options || {}))
  const res = await fetch('/api/plan', { method: 'POST', body: fd })
  if (!res.ok) throw new Error((await res.json()).error || 'Plan failed')
  return res.json()
}

async function apiStartExportJob({ plan, options }) {
  const res = await fetch('/api/export_job', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ plan, options })
  })
  if (!res.ok) throw new Error((await res.json()).error || 'Export failed')
  return res.json()
}

async function apiDownloadExportJobPptx(jobId) {
  const res = await fetch(`/api/export_job/${jobId}/pptx`)
  if (!res.ok) throw new Error((await res.json()).error || 'Download failed')
  const blob = await res.blob()
  const cd = res.headers.get('content-disposition') || ''
  const match = cd.match(/filename="([^"]+)"/)
  const filename = match?.[1] || 'Deck.pptx'
  return { blob, filename }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}

function thumbBg(kind = '') {
  const k = (kind || '').toString().toLowerCase()
  if (k.includes('title') || k.includes('cover') || k.includes('hero')) return 'bg-gradient-to-br from-fuchsia-200 via-white to-sky-200'
  if (k.includes('creative') || k.includes('concept') || k.includes('big_idea')) return 'bg-gradient-to-br from-amber-200 via-white to-pink-200'
  if (k.includes('visual')) return 'bg-gradient-to-br from-emerald-200 via-white to-sky-200'
  if (k.includes('challenge') || k.includes('problem')) return 'bg-gradient-to-br from-rose-200 via-white to-orange-200'
  if (k.includes('opportunity') || k.includes('insight')) return 'bg-gradient-to-br from-lime-200 via-white to-emerald-200'
  if (k.includes('execution')) return 'bg-gradient-to-br from-violet-200 via-white to-fuchsia-200'
  if (k.includes('timeline') || k.includes('roadmap')) return 'bg-gradient-to-br from-sky-200 via-white to-indigo-200'
  if (k.includes('kpi') || k.includes('dashboard') || k.includes('qbr')) return 'bg-gradient-to-br from-emerald-200 via-white to-amber-200'
  return 'bg-gradient-to-br from-sky-100 via-white to-pink-100'
}

export default function App() {
  const [file, setFile] = useState(null)
  const [extraText, setExtraText] = useState('')
  const [options, setOptions] = useState({
    provider: 'openai',
    deckType: '',
    nSlides: 10,
    deckStyle: 'classic',
    voiceProfile: 'witty_agency',
    vibe: 'Modern, premium, ad-agency-level',
    audience: 'general',
    language: 'English',
    imageStyle: 'Editorial photography + clean gradients, premium, high contrast'
  })

  const [status, setStatus] = useState('')
  const [planning, setPlanning] = useState(false)
  const [planPct, setPlanPct] = useState(0)
  const [extractedText, setExtractedText] = useState('')
  const [plan, setPlan] = useState(null)

  const [exportJobId, setExportJobId] = useState(null)
  const [exportRunning, setExportRunning] = useState(false)
  const [exportThumbs, setExportThumbs] = useState([]) // array of PNG URLs (or null while pending)
  const [exportPhase, setExportPhase] = useState('')
  const [exportFilename, setExportFilename] = useState('Deck.pptx')

  const canPlan = useMemo(()=> Boolean(file || extraText.trim()), [file, extraText])

  useEffect(() => {
    if (!planning) { setPlanPct(0); return }
    const start = Date.now()
    setPlanPct(0)
    const t = setInterval(() => {
      const elapsed = Date.now() - start
      // Ease to 92% while waiting for the API response
      const pct = Math.min(92, Math.round((elapsed / 6500) * 92))
      setPlanPct(pct)
    }, 120)
    return () => clearInterval(t)
  }, [planning])

  const onGenerateOutline = async () => {
    setPlanning(true)
    setStatus('Generating outline…')
    try {
      const data = await apiPlan({ file, text: extraText, options })
      setExtractedText(data.extractedText || '')
      setPlan(data.plan)
      setStatus('Outline ready. Edit it, then export.')
    } catch (e) {
      setStatus(e.message)
    } finally {
      setPlanning(false)
    }
  }

  const onExport = async () => {
    if (!plan) return
    setExportRunning(true)
    setExportJobId(null)
    setExportPhase('Starting export…')
    setExportFilename(`${(plan.deck_title || 'Deck').replace(/[^a-zA-Z0-9._-]/g,'_')}.pptx`)
    const total = Array.isArray(plan.slides) ? plan.slides.length : 0
    setExportThumbs(Array.from({ length: total }, () => null))
    setStatus('Export started…')
    try {
      const started = await apiStartExportJob({ plan, options })
      const jobId = started.jobId
      setExportJobId(jobId)
      if (started.filenameHint) setExportFilename(started.filenameHint)
      if (Number.isFinite(started.totalSlides) && started.totalSlides > 0) {
        setExportThumbs(Array.from({ length: started.totalSlides }, () => null))
      }

      const es = new EventSource(`/api/export_job/${jobId}/stream`)

      es.addEventListener('meta', (ev) => {
        try {
          const d = JSON.parse(ev.data)
          if (d?.filename) setExportFilename(d.filename)
          if (Number.isFinite(d?.totalSlides) && d.totalSlides > 0) {
            setExportThumbs((prev) =>
              prev.length === d.totalSlides ? prev : Array.from({ length: d.totalSlides }, () => null)
            )
          }
        } catch {}
      })

      es.addEventListener('status', (ev) => {
        try {
          const d = JSON.parse(ev.data)
          setExportPhase(d?.message || d?.phase || '')
        } catch {}
      })

      es.addEventListener('thumbnail', (ev) => {
        try {
          const d = JSON.parse(ev.data)
          const idx = d?.index
          const url = d?.url
          if (!Number.isFinite(idx) || !url) return
          setExportThumbs((prev) => {
            const next = [...prev]
            next[idx] = url
            return next
          })
        } catch {}
      })

      es.addEventListener('done', async (ev) => {
        try {
          const d = JSON.parse(ev.data)
          if (d?.filename) setExportFilename(d.filename)
        } catch {}

        try {
          setExportPhase('Downloading PPTX…')
          const { blob, filename } = await apiDownloadExportJobPptx(jobId)
          downloadBlob(blob, filename)
          setStatus('Done. PPTX downloaded.')
        } catch (e) {
          setStatus(e.message)
        } finally {
          es.close()
          setExportRunning(false)
          setExportPhase('')
        }
      })

      es.addEventListener('error', (ev) => {
        // Note: EventSource uses 'error' for network issues too.
        try {
          const d = ev?.data ? JSON.parse(ev.data) : null
          if (d?.message) setStatus(d.message)
        } catch {}
      })
    } catch (e) {
      setStatus(e.message)
      setExportRunning(false)
      setExportPhase('')
    }
  }


  const exportPct = useMemo(() => {
    const doneCount = exportThumbs.filter(Boolean).length
    const totalCount = exportThumbs.length || 0
    if (!totalCount) return 0
    return Math.min(100, Math.round((doneCount / totalCount) * 100))
  }, [exportThumbs])

  const overlayPct = planning ? planPct : exportPct


  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_20%_10%,rgba(59,130,246,0.18),transparent_45%),radial-gradient(circle_at_80%_15%,rgba(236,72,153,0.18),transparent_45%),radial-gradient(circle_at_40%_90%,rgba(16,185,129,0.18),transparent_45%),linear-gradient(180deg,rgba(255,255,255,1),rgba(255,255,255,0.7))]">
      {(planning || exportRunning) && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-white/30 backdrop-blur-md">
          <div className="w-[min(560px,92vw)] rounded-3xl border border-white/50 bg-white/55 p-6 shadow-[0_20px_80px_rgba(0,0,0,0.12)] backdrop-blur-2xl ring-1 ring-white/40">
            <div className="text-sm font-semibold text-zinc-800">{planning ? 'Generating outline' : 'Exporting PPTX'}</div>
            <div className="mt-1 text-xs text-zinc-600">
              {planning ? 'Turning your brief into a clean slide plan…' : (exportPhase || 'Rendering slides…')}
            </div>

            <div className="mt-4">
              <div className="flex items-center justify-between text-[11px] font-semibold text-zinc-700">
                <span>{planning ? 'Planning' : `${exportThumbs.filter(Boolean).length}/${exportThumbs.length || 0} slides`}</span>
                <span>{overlayPct}%</span>
              </div>
              <div className="mt-2 h-2.5 w-full overflow-hidden rounded-full bg-white/55">
                <div
                  className="h-2.5 bg-gradient-to-r from-sky-400 via-fuchsia-400 to-amber-300"
                  style={{ width: `${overlayPct}%` }}
                />
              </div>
            </div>

            {!planning && (
              <div className="mt-4 text-[11px] text-zinc-500">
                Thumbnails render live during export. Keep this tab open.
              </div>
            )}
          </div>
        </div>
      )}

      <header className="mx-auto max-w-6xl px-5 pt-10 pb-6">
        <div className="flex items-end justify-between gap-4">
          <div>
            <div className="text-xs font-semibold text-fuchsia-600">DeckSmith Pro</div>
            <h1 className="mt-1 text-3xl font-black tracking-tight md:text-4xl">DeckSmith Pro</h1>
            <p className="mt-2 max-w-2xl text-sm text-zinc-600">
              Turn a brief into a bold, editable PowerPoint — with slide-by-slide control and export-ready layouts.
            </p>
          </div>
          <a
            className="hidden rounded-xl border border-white/60 bg-white/55 px-4 py-2 text-sm font-semibold text-zinc-700 hover:bg-white/80 md:inline-flex"
            href="https://platform.openai.com/docs"
            target="_blank"
            rel="noreferrer"
          >
            OpenAI Docs
          </a>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-5 pb-14">
        <div className="grid gap-6 lg:grid-cols-2">
          <div className="space-y-6">
            <Dropzone file={file} setFile={setFile} />

            <div className="rounded-2xl border border-white/50 bg-white/45 p-5 shadow-sm backdrop-blur-2xl ring-1 ring-white/40">
              <div className="text-sm font-semibold">Optional extra notes</div>
              <textarea
                value={extraText}
                onChange={(e)=>setExtraText(e.target.value)}
                className="mt-2 h-32 w-full rounded-2xl border border-white/60 bg-white/55 p-3 text-sm outline-none focus:border-zinc-400"
                placeholder="Paste any extra context, messaging, constraints, tone…"
              />
            
            </div>

            <div className="rounded-2xl border border-white/50 bg-white/45 p-5 shadow-[0_10px_40px_rgba(0,0,0,0.06)] backdrop-blur-2xl ring-1 ring-white/40">
              <div className="flex flex-wrap items-center gap-3">
                <button
                  disabled={!canPlan || planning || exportRunning}
                  onClick={onGenerateOutline}
                  className={`rounded-xl px-4 py-2 text-sm font-semibold text-white ${
                    canPlan ? 'bg-gradient-to-r from-pink-500 to-amber-400 hover:from-pink-600 hover:to-amber-500' : 'bg-zinc-300'
                  }`}
                >
                  Generate outline
                </button>

                <button
                  disabled={!plan || exportRunning || planning}
                  onClick={onExport}
                  className={`rounded-xl px-4 py-2 text-sm font-semibold ${
                    plan ? 'bg-white/70 hover:bg-white text-zinc-900 border border-white/70' : 'bg-zinc-100/60 text-zinc-400 border border-white/60'
                  }`}
                >
                  Export PPTX
                </button>

                <div className="text-xs text-zinc-600">{status}</div>
              </div>
              <div className="mt-3 text-[11px] text-zinc-500">
                Export thumbnails appear only during export.
              </div>
            </div>

            <div className="rounded-2xl border border-white/50 bg-white/45 p-5 shadow-sm backdrop-blur-2xl ring-1 ring-white/40">
              <div className="text-sm font-semibold">Generation settings</div>
              <div className="mt-4 grid gap-4 md:grid-cols-2">
                <label className="block md:col-span-2">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">AI provider</div>
                  <select
                    value={options.provider}
                    onChange={(e)=>setOptions(o=>({...o, provider: e.target.value}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  >
                    <option value="openai">ChatGPT + DALL·E (OpenAI)</option>
                    <option value="gemini">Gemini (Nano Banana Pro)</option>
                  </select>
                  <div className="mt-1 text-[11px] text-zinc-500">
                    Gemini requires <span className="font-mono">GEMINI_API_KEY</span> set in <span className="font-mono">server/.env</span>.
                  </div>
                </label>
                <label className="block">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Slides</div>
                  <input
                    type="number"
                    min={5}
                    max={30}
                    value={options.nSlides}
                    disabled={options.deckType === 'ad_agency'}
                    onChange={(e)=>setOptions(o=>({...o, nSlides: Number(e.target.value)}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  />
                  {options.deckType === 'ad_agency' && (
                    <div className="mt-1 text-[11px] text-zinc-500">
                      Agency creative decks are locked to a 10‑slide flow.
                    </div>
                  )}
                </label>

                <label className="block">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Presentation type</div>
                  <select
                    value={options.deckType}
                    onChange={(e)=>{
                      const v = e.target.value
                      setOptions(o=>({
                        ...o,
                        deckType: v,
                        nSlides: v === 'ad_agency' ? 10 : o.nSlides
                      }))
                    }}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  >
                    <option value="">Auto-detect</option>
                    <option value="investor_pitch">Investor pitch deck</option>
                    <option value="sales_deck">Sales deck</option>
                    <option value="business_proposal">Business proposal</option>
                    <option value="marketing_strategy">Marketing strategy & campaign</option>
                    <option value="qbr">Quarterly business review (QBR)</option>
                    <option value="product_roadmap">Product roadmap</option>
                    <option value="company_profile">Company profile / credentials</option>
                    <option value="training_workshop">Training / workshop</option>
                    <option value="project_status_update">Project status update</option>
                    <option value="keynote_thought_leadership">Keynote / thought leadership</option>
                    <option value="ad_agency">Creative / agency deck</option>
                  </select>
                </label>

                <label className="block">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Copy voice</div>
                  <select
                    value={options.voiceProfile}
                    onChange={(e)=>setOptions(o=>({...o, voiceProfile: e.target.value}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  >
                    <option value="witty_agency">Witty agency (default)</option>
                    <option value="cinematic_minimal">Cinematic minimal</option>
                    <option value="corporate_clear">Corporate clear</option>
                    <option value="academic_formal">Academic formal</option>
                  </select>
                  <div className="mt-1 text-[11px] text-zinc-500">
                    Controls copy rules (headlines as claims, section pacing, callbacks). Visual style is set by Deck style.
                  </div>
                </label>

                <label className="block">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Language</div>
                  <input
                    value={options.language}
                    onChange={(e)=>setOptions(o=>({...o, language: e.target.value}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  />
                </label>
                <label className="block md:col-span-2">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Vibe</div>
                  <input
                    value={options.vibe}
                    onChange={(e)=>setOptions(o=>({...o, vibe: e.target.value}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  />
                </label>
                <label className="block md:col-span-2">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Deck style</div>
                  <select
                    value={options.deckStyle}
                    onChange={(e)=>setOptions(o=>({...o, deckStyle: e.target.value}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  >
                    <option value="classic">Classic Premium</option>
                    <option value="agency_typographic">Agency Typographic</option>
                    <option value="neo_brutal">Neo‑Brutalist</option>
                    <option value="bento_minimal">Bento Minimal</option>
                    <option value="gradient_mesh">Gradient Mesh</option>
                  </select>
                  <div className="mt-1 text-[11px] text-zinc-500">
                    Switches PPTX typography + layout chrome. You can still override the image look below.
                  </div>
                </label>

                <label className="block md:col-span-2">
                  <div className="mb-1 text-xs font-semibold text-zinc-600">Image style</div>
                  <input
                    value={options.imageStyle}
                    onChange={(e)=>setOptions(o=>({...o, imageStyle: e.target.value}))}
                    className="w-full rounded-xl border border-white/60 bg-white/55 px-3 py-2 text-sm"
                  />
                </label>
              </div>

              <div className="mt-5 flex flex-wrap items-center gap-3">
                <button
                  disabled={!canPlan || planning}
                  onClick={onGenerateOutline}
                  className={`rounded-xl px-4 py-2 text-sm font-semibold text-white ${
                    canPlan ? 'bg-gradient-to-r from-pink-500 to-amber-400 hover:from-pink-600 hover:to-amber-500' : 'bg-zinc-300'
                  }`}
                >
                  Generate outline
                </button>

                <button
                  disabled={!plan || exportRunning}
                  onClick={onExport}
                  className={`rounded-xl px-4 py-2 text-sm font-semibold ${
                    plan ? 'bg-white hover:bg-amber-50 text-zinc-900 border border-amber-200' : 'bg-zinc-100 text-zinc-400 border border-white/60'
                  }`}
                >
                  Export PPTX
                </button>

                <div className="text-xs text-zinc-600">{status}</div>
              </div>

              {(planning || exportRunning) && (
                <div className="mt-3 h-2 w-full overflow-hidden rounded-full bg-white/55">
                  <div className={`h-2 ${planning ? 'w-2/3 animate-pulse' : 'w-full'} bg-gradient-to-r from-sky-400 via-fuchsia-400 to-amber-300`} />
                </div>
              )}
            </div>

            {extractedText && (
              <div className="rounded-2xl border border-white/50 bg-white/45 p-5 shadow-sm backdrop-blur-2xl ring-1 ring-white/40">
                <div className="flex items-center justify-between">
                  <div className="text-sm font-semibold">Extracted text (read-only)</div>
                  <button
                    onClick={()=>navigator.clipboard.writeText(extractedText)}
                    className="rounded-xl border border-white/60 px-3 py-2 text-xs hover:bg-white/80"
                  >
                    Copy
                  </button>
                </div>
                <pre className="mt-3 max-h-56 overflow-auto whitespace-pre-wrap rounded-2xl bg-white/50 p-3 text-xs text-zinc-700">
                  {extractedText}
                </pre>
              </div>
            )}
          </div>

          <div>
            {!plan ? (
              <div className="rounded-2xl border border-white/50 bg-white/45 p-8 shadow-sm backdrop-blur-2xl ring-1 ring-white/40">
                <div className="text-sm font-semibold">Your outline will appear here</div>
                <p className="mt-2 text-sm text-zinc-600">
                  Upload a brief (PDF/DOC/DOCX/PPTX) or paste notes, then click <span className="font-semibold">Generate outline</span>.
                  You’ll be able to edit every slide before exporting.
                </p>
                <div className="mt-6 rounded-2xl bg-white/50 p-4 text-xs text-zinc-600">
                  <div className="font-semibold text-zinc-700">Pro tip</div>
                  <div className="mt-1">
                    Best results come from briefs that include: objective, audience, key messages, proof points, and desired tone.
                  </div>
                </div>
              </div>
            ) : (
              <div className="space-y-6">
                {exportRunning && (
                  <div className="rounded-2xl border border-white/60 bg-white/55 p-5 shadow-sm backdrop-blur-xl">
                    <div className="flex items-end justify-between gap-4">
                      <div>
                        <div className="text-sm font-semibold">Exporting deck</div>
                        <div className="mt-1 text-xs text-zinc-600">
                          {exportPhase || 'Rendering slides…'} {exportJobId ? <span className="font-mono text-[10px] text-zinc-400">({exportJobId})</span> : null}
                        </div>
                      </div>
                      <div className="text-xs font-semibold text-zinc-600">{exportThumbs.length || 0} slides</div>
                    </div>

                    {(() => {
                      const doneCount = exportThumbs.filter(Boolean).length
                      const totalCount = exportThumbs.length || 1
                      const pct = Math.round((doneCount / totalCount) * 100)
                      return (
                        <div className="mt-4">
                          <div className="flex items-center justify-between text-[11px] font-semibold text-zinc-600">
                            <span>{doneCount}/{totalCount} slides rendered</span>
                            <span>{pct}%</span>
                          </div>
                          <div className="mt-2 h-2 w-full overflow-hidden rounded-full bg-white/55">
                            <div className="h-2 bg-gradient-to-r from-emerald-400 via-sky-400 to-fuchsia-400" style={{ width: `${pct}%` }} />
                          </div>
                        </div>
                      )
                    })()}

                    <div className="mt-4 flex gap-3 overflow-x-auto pb-2">
                      {(exportThumbs.length ? exportThumbs : Array.from({ length: 6 }, () => null)).map((url, idx) => {
                        const meta = Array.isArray(plan?.slides) ? plan.slides[idx] : null
                        const kind = meta?.kind || meta?.layout || ''
                        const title = meta?.title || (meta ? 'Untitled' : '')
                        return (
                          <div key={idx} className="w-56 shrink-0 overflow-hidden rounded-2xl border border-white/70 bg-white/70 shadow-sm">
                            <div className="relative h-28 overflow-hidden">
                              {url ? (
                                <img src={url} alt={`Slide ${idx + 1}`} className="h-28 w-full object-cover" />
                              ) : (
                                <div className="h-28 w-full animate-pulse bg-gradient-to-br from-zinc-100 via-white to-zinc-200" />
                              )}
                              <div className="absolute inset-0 bg-gradient-to-t from-black/35 via-black/0 to-black/0" />
                              <div className="absolute left-3 top-3 rounded-lg bg-white/70 px-2 py-1 text-[11px] font-semibold text-zinc-700 backdrop-blur">
                                Slide {idx + 1}
                              </div>
                            </div>
                            <div className="p-3">
                              <div className="text-[11px] font-semibold text-zinc-500">{kind || 'content'}</div>
                              <div className="mt-1 line-clamp-2 text-xs font-semibold text-zinc-800">{title || '—'}</div>
                              {!url && (
                                <div className="mt-2 text-[11px] text-zinc-500">Rendering…</div>
                              )}
                              {url && (
                                <a href={url} target="_blank" rel="noreferrer" className="mt-2 inline-block text-[11px] font-semibold text-fuchsia-700 hover:underline">
                                  Open thumbnail
                                </a>
                              )}
                            </div>
                          </div>
                        )
                      })}                    </div>

                    <div className="mt-3 text-[11px] text-zinc-500">
                      Thumbnails render live during export (LibreOffice PNGs).
                    </div>
                  </div>
                )}

                <PlanEditor plan={plan} setPlan={setPlan} />
              </div>
            )}
          </div>
        </div>

        <footer className="mt-10 text-xs text-zinc-500">
          API keys must stay server-side. Don’t commit <span className="font-mono">.env</span>.
        </footer>
      </main>
    </div>
  )
}