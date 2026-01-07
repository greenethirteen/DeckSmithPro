# Provider dropdown patch (OpenAI vs Gemini)

## What this adds
- A new **AI provider** dropdown in the UI:
  - **ChatGPT + DALL·E (OpenAI)**
  - **Gemini (Nano Banana Pro)**
- The selected provider is sent to the server as `options.provider`.
- The server routes:
  - **Planning** (JSON outline) to OpenAI or Gemini.
  - **Image generation** to OpenAI or Gemini (Nano Banana Pro uses `gemini-3-pro-image-preview`).

## How to apply
From your project root (the folder that contains `client/` and `server/`):

```bash
# Option 1: copy files over manually from this patch zip (recommended)
# Option 2: if you unzip this patch on top of your project root, it will overwrite the right files.
```

## Env vars
OpenAI (existing):
- `OPENAI_API_KEY`

Gemini (new; required only if provider = gemini):
- `GEMINI_API_KEY`
- Optional overrides:
  - `GEMINI_TEXT_MODEL` (default: `gemini-3-pro`)
  - `GEMINI_IMAGE_MODEL` (default: `gemini-3-pro-image-preview`)

## Notes
- Image prompts now explicitly forbid any text/typography/words so you don’t get background text that clashes with editable PPTX text.
- Gemini image generation requests a 16:9 aspect ratio; PPTX export still uses smart cover-cropping (no squeezing).
