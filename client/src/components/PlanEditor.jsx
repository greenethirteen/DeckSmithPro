import React, { useMemo, useState } from 'react'

function safeArray(v) {
  return Array.isArray(v) ? v : v ? [String(v)] : []
}

function thumbClass(kind) {
  const k = (kind || '').toString().toLowerCase()
  if (k.includes('title') || k.includes('cover')) return 'bg-gradient-to-br from-pink-200 via-white to-amber-200'
  if (k.includes('creative') || k.includes('concept') || k.includes('big_idea')) return 'bg-gradient-to-br from-amber-200 via-white to-fuchsia-200'
  if (k.includes('visual')) return 'bg-gradient-to-br from-sky-200 via-white to-emerald-200'
  if (k.includes('challenge') || k.includes('problem')) return 'bg-gradient-to-br from-rose-200 via-white to-orange-200'
  if (k.includes('opportunity') || k.includes('insight')) return 'bg-gradient-to-br from-emerald-200 via-white to-lime-200'
  if (k.includes('execution')) return 'bg-gradient-to-br from-violet-200 via-white to-pink-200'
  return 'bg-gradient-to-br from-amber-100 via-white to-pink-100'
}

export default function PlanEditor({ plan, setPlan }) {
  const [view, setView] = useState('slides') // slides | json
  const slides = useMemo(() => safeArray(plan?.slides), [plan])

  if (!plan) {
    return (
      <div className="rounded-2xl border border-white/60 bg-white/60 p-5 shadow-sm backdrop-blur-xl">
        <div className="text-sm font-semibold">Outline</div>
        <div className="mt-2 text-sm text-zinc-500">Generate an outline to start editing.</div>
      </div>
    )
  }

  const updateSlide = (idx, patch) => {
    setPlan((p) => {
      const next = { ...(p || {}) }
      next.slides = safeArray(next.slides).map((s, i) => (i === idx ? { ...(s || {}), ...patch } : s))
      return next
    })
  }

  const scrollToSlide = (idx) => {
    const el = document.getElementById(`slide-${idx}`)
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  return (
    <div className="space-y-6">
      {/* Outline editor */}
      <div className="rounded-2xl border border-white/60 bg-white/60 p-5 shadow-sm backdrop-blur-xl">
        <div className="flex items-center justify-between gap-3">
          <div>
            <div className="text-sm font-semibold">Outline editor</div>
            <div className="mt-1 text-xs text-zinc-500">Edit slide copy before exporting.</div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={() => setView('slides')}
              className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${
                view === 'slides'
                  ? 'bg-gradient-to-r from-fuchsia-500 to-sky-400 text-white'
                  : 'border border-white/60 bg-white text-zinc-700 hover:bg-white/70'
              }`}
            >
              Slides
            </button>
            <button
              onClick={() => setView('json')}
              className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${
                view === 'json'
                  ? 'bg-gradient-to-r from-fuchsia-500 to-sky-400 text-white'
                  : 'border border-white/60 bg-white text-zinc-700 hover:bg-white/70'
              }`}
            >
              JSON
            </button>
          </div>
        </div>

        {view === 'json' ? (
          <textarea
            className="mt-4 h-[520px] w-full rounded-2xl border border-white/60 bg-white p-3 font-mono text-[12px] outline-none focus:border-white/80"
            value={JSON.stringify(plan, null, 2)}
            onChange={(e) => {
              try {
                const parsed = JSON.parse(e.target.value)
                setPlan(parsed)
              } catch {
                // ignore until valid
              }
            }}
          />
        ) : (
          <div className="mt-4 space-y-4">
            {slides.map((s, idx) => {
              const type = (s?.kind || s?.type || s?.layout || 'content').toString()
              const title = (s?.title || '').toString()
              const subtitle = (s?.subtitle || '').toString()
              const bullets = safeArray(s?.bullets)
              const imagePrompt = (s?.image_prompt || '').toString()

              return (
                <div key={idx} id={`slide-${idx}`} className="rounded-2xl border border-white/60 p-4">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <div className="text-xs font-semibold text-zinc-500">Slide {idx + 1} • {type}</div>
                      <div className="mt-1 text-sm font-semibold text-zinc-900">{title || 'Untitled'}</div>
                    </div>
                  </div>

                  <div className="mt-3 grid gap-3 md:grid-cols-2">
                    <label className="block md:col-span-2">
                      <div className="mb-1 text-xs font-semibold text-zinc-600">Title</div>
                      <input
                        value={title}
                        onChange={(e) => updateSlide(idx, { title: e.target.value })}
                        className="w-full rounded-xl border border-white/60 bg-white px-3 py-2 text-sm"
                      />
                    </label>
                    <label className="block md:col-span-2">
                      <div className="mb-1 text-xs font-semibold text-zinc-600">Subtitle / body</div>
                      <textarea
                        value={subtitle}
                        onChange={(e) => updateSlide(idx, { subtitle: e.target.value })}
                        className="h-20 w-full rounded-xl border border-white/60 bg-white px-3 py-2 text-sm"
                      />
                    </label>
                    <label className="block md:col-span-2">
                      <div className="mb-1 text-xs font-semibold text-zinc-600">Bullets (one per line)</div>
                      <textarea
                        value={bullets.join('\n')}
                        onChange={(e) => updateSlide(idx, { bullets: e.target.value.split(/\r?\n/).map(v=>v.trim()).filter(Boolean) })}
                        className="h-28 w-full rounded-xl border border-white/60 bg-white px-3 py-2 text-sm"
                        placeholder="• ..."
                      />
                    </label>
                    <label className="block md:col-span-2">
                      <div className="flex items-center justify-between">
                        <div className="mb-1 text-xs font-semibold text-zinc-600">Background image prompt</div>
                        <div className="flex items-center gap-2 text-[11px] font-semibold text-zinc-500">
                          <button
                            type="button"
                            onClick={() => updateSlide(idx, { image_prompt: 'NONE' })}
                            className="rounded-md border border-white/70 bg-white px-2 py-0.5 hover:bg-white/70"
                          >
                            Use none
                          </button>
                          <button
                            type="button"
                            onClick={() => updateSlide(idx, { image_prompt: '' })}
                            className="rounded-md border border-white/70 bg-white px-2 py-0.5 hover:bg-white/70"
                          >
                            Clear
                          </button>
                        </div>
                      </div>
                      <textarea
                        value={imagePrompt}
                        onChange={(e) => updateSlide(idx, { image_prompt: e.target.value })}
                        className="h-24 w-full rounded-xl border border-white/60 bg-white px-3 py-2 text-sm"
                        placeholder="Describe the background image you want for this slide."
                      />
                      <div className="mt-1 text-[11px] text-zinc-500">
                        This prompt replaces the background image for this slide on export.
                      </div>
                    </label>
                  </div>
                </div>
              )
            })}
          </div>
        )}
      </div>
    </div>
  )
}
