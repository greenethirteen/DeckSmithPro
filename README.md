# DeckSmith Pro (v1)

## Run (single terminal)

```bash
cd decksmith-pro-v1
cp server/.env.example server/.env
# set OPENAI_API_KEY and/or GEMINI_API_KEY
npm install
npm run dev
```

- Client: http://localhost:5173
- Server: http://localhost:8787/api/health

## Run (two terminals)

Server:
```bash
npm --prefix server install
npm --prefix server run dev
```

Client:
```bash
npm --prefix client install
npm --prefix client run dev
```

## Graphics renderer (optional)

Environment toggles:
- `PPTX_GRAPHICS_PREFER_SVG` (default `true`)
- `PPTX_GRAPHICS_FORCE_PNG` (default `false`)
- `ICONIFY_API_BASE` (default `https://api.iconify.design`)
- `MERMAID_CLI_PATH` (optional path to `mmdc` if not in `node_modules/.bin`)
