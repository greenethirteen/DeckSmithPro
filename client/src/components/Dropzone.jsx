import React, { useCallback, useRef, useState } from 'react'

export default function Dropzone({ file, setFile }) {
  const inputRef = useRef(null)
  const [dragOver, setDragOver] = useState(false)

  const onPick = useCallback(() => inputRef.current?.click(), [])

  const onFiles = useCallback(
    (files) => {
      const f = files?.[0]
      if (!f) return
      const name = (f.name || '').toLowerCase();
      const ok = ['.pdf', '.doc', '.docx', '.pptx'].some(ext => name.endsWith(ext));
      if (!ok) {
        alert('Please choose a PDF, DOC/DOCX, or PPTX brief.');
        return;
      }
      setFile(f)
    },
    [setFile]
  )

  return (
    <div
      className={`rounded-2xl border bg-white/60 p-5 shadow-sm backdrop-blur-xl transition ${
        dragOver ? 'border-white/80' : 'border-white/60'
      }`}
      onDragEnter={(e) => {
        e.preventDefault()
        setDragOver(true)
      }}
      onDragOver={(e) => {
        e.preventDefault()
        setDragOver(true)
      }}
      onDragLeave={(e) => {
        e.preventDefault()
        setDragOver(false)
      }}
      onDrop={(e) => {
        e.preventDefault()
        setDragOver(false)
        onFiles(Array.from(e.dataTransfer.files || []))
      }}
    >
      <div className="flex items-center justify-between gap-3">
        <div>
          <div className="text-sm font-semibold">Upload brief file</div>
          <div className="mt-1 text-xs text-zinc-600">
            Drag & drop a PDF, DOC/DOCX, or PPTX — or click to select.
          </div>
        </div>
        <button
          onClick={onPick}
          className="rounded-xl border border-white/60 bg-white px-4 py-2 text-sm font-semibold text-zinc-800 hover:bg-white/50"
        >
          Choose file
        </button>
      </div>

      <input
        ref={inputRef}
        type="file"
        accept=".pdf,.doc,.docx,.pptx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-word,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        className="hidden"
        onChange={(e) => onFiles(Array.from(e.target.files || []))}
      />

      {file && (
        <div className="mt-4 flex items-center justify-between rounded-xl border border-white/60 bg-white/50 px-3 py-2">
          <div className="min-w-0">
            <div className="truncate text-sm font-semibold text-zinc-800">{file.name}</div>
            <div className="text-xs text-zinc-600">{Math.round(file.size / 1024)} KB</div>
          </div>
          <button
            onClick={() => setFile(null)}
            className="rounded-lg px-2 py-1 text-xs font-semibold text-zinc-600 hover:bg-white"
          >
            Remove
          </button>
        </div>
      )}
    </div>
  )
}
