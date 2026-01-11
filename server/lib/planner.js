import { getOpenAIClient, getModels } from './openai.js';
import { geminiGenerateJson } from './gemini.js';

const SAFE_MAX_SLIDES = 30;

// Copy voice profiles ("agency brain" constraints live here).
// These are *writing* constraints (headlines, transitions, diction), not visual style.
export const VOICE_PROFILES = [
  'witty_agency',
  'cinematic_minimal',
  'corporate_clear',
  'academic_formal'
];

const VOICE_RULES = {
  witty_agency: {
    name: 'Witty Agency',
    tagline: 'Clever, confident, punchy. Feels like a CD wrote it the night before the client pitch — tight, not try-hard.',
    headline_rules: [
      'Every slide headline must be a declarative claim (no labels like “The Challenge”, “Objectives”, “Strategy”).',
      'One idea per headline. Avoid “and”.',
      'Keep headlines short (3–9 words), high-contrast, a little surprising.',
      'Prefer contrast and reframes (e.g., “Not X. Y.” / “From __ to __”).'
    ],
    subhead_rules: [
      'Subhead (subtitle) explains the headline in one sentence. No new concept introduced here.',
      'If you need more context, put it in bullets or speaker notes, not a second thesis.'
    ],
    section_rules: [
      'Each narrative section introduces ONE new concept only. Everything else in the section is proof, implication, or example.',
      'Every section must end with a bridge line that tees up the next section.',
      'Callbacks: pick 1–3 anchor lines (platform phrases) and repeat them verbatim later (no paraphrases).'
    ],
    diction_rules: [
      'Avoid corporate filler and buzzwords. No “synergy”, “leverage”, “stakeholders”, “robust”, “seamless”, “omnichannel” unless the brief explicitly wants corporate tone.',
      'No cringe jokes, no slang, no memes. Witty = smart and minimal.',
      'Use specific nouns/verbs; limit adjectives.'
    ],
    forbidden_terms: ['synergy', 'leverage', 'stakeholders', 'robust', 'seamless', 'omnichannel', 'best-in-class', 'world-class'],
  },
  cinematic_minimal: {
    name: 'Cinematic Minimal',
    tagline: 'Sparse, visual, emotionally driven. Big claims, few words.',
    headline_rules: ['3–7 words, declarative, cinematic.'],
    subhead_rules: ['Optional. If used, 6–12 words max.'],
    section_rules: ['One new concept per section. Strong bridges.'],
    diction_rules: ['No jargon. No jokes.'],
    forbidden_terms: []
  },
  corporate_clear: {
    name: 'Corporate Clear',
    tagline: 'Executive clarity, structured, direct. Less personality, more proof.',
    headline_rules: ['Declarative, specific, no hype.'],
    subhead_rules: ['Explain claim, quantify where possible.'],
    section_rules: ['One new concept per section. Strong bridges.'],
    diction_rules: ['Avoid fluff.'],
    forbidden_terms: ['game-changer', 'revolutionary']
  },
  academic_formal: {
    name: 'Academic Formal',
    tagline: 'Neutral tone, careful claims, definitions first.',
    headline_rules: ['Declarative, precise, no punchlines.'],
    subhead_rules: ['Define terms, state assumptions.'],
    section_rules: ['One new concept per section. Clear bridges.'],
    diction_rules: ['No jokes. No hype.'],
    forbidden_terms: []
  }
};

// High-level deck intents your system can plan for.
export const DECK_TYPES = [
  'ad_agency',
  'investor_pitch',
  'sales_deck',
  'business_proposal',
  'marketing_strategy',
  'qbr',
  'product_roadmap',
  'company_profile',
  'training_workshop',
  'project_status_update',
  'keynote_thought_leadership',
  'other'
];

// Renderer-supported layouts.
export const SLIDE_LAYOUTS = [
  // existing
  'hero',
  'split',
  'full_bleed',
  'quote',
  'stats',
  'two_column',
  // agency-only (bold typographic)
  'agency_center',
  'agency_half',
  'agency_infographic',
  // new (reusable across many deck types)
  'section_header',
  'agenda',
  'cards',
  'image_caption',
  'timeline',
  'kpi_dashboard',
  'traffic_light',
  'table',
  'pricing',
  'comparison_matrix',
  'process_steps',
  'team_grid',
  'logo_wall',
  'cta',
  // additional business slide types
  'swot',
  'funnel',
  'now_next_later',
  'okr',
  'case_study',
  'chart_bar',
  'chart_line',
  'diagram',
  'org_chart',
  'faq',
  'appendix',
  'infographic_3'
];

// Narrative templates (story-first), not visual templates.
// These are used by the narrative step to enforce slide-to-slide flow.
const DECK_RECIPES = {
  // Fixed 10-slide agency creative story (what you requested):
  // Title → Current Audience → About the Brand → Challenge → Opportunity → Communication Pillars/Media Types → Creative Concept → Visual Identity → Execution Example → Thank You
  agency_creative_10: [
    'title',
    'current_audience',
    'about_brand',
    'challenge',
    'opportunity',
    'communication_pillars',
    'creative_concept',
    'visual_identity',
    'execution_example',
    'thank_you'
  ],
  // The exact arc your MDLBEAST case study reference follows:
  // proof of signal → align to brand → articulate challenge → reframe to opportunity → pillars → platform → visual system → executions → objectives → next steps → close
  marketing_case_study: [
    'cover',
    'proof_or_current_state',
    'brand_alignment',
    'challenge',
    'insight_or_reframe',
    'strategy_pillars',
    'creative_platform',
    'visual_direction',
    'execution_examples',
    'rollout_or_system',
    'measurement_objectives',
    'next_steps',
    'close'
  ],
  investor_pitch: ['cover', 'problem', 'solution', 'market', 'product', 'business_model', 'traction', 'go_to_market', 'competition', 'team', 'ask', 'close'],
  sales_deck: ['cover', 'customer_pain', 'why_now', 'solution', 'how_it_works', 'benefits', 'proof', 'pricing', 'implementation', 'cta', 'close'],
  business_proposal: ['cover', 'context', 'objectives', 'scope', 'approach', 'deliverables', 'timeline', 'investment', 'team', 'risks', 'next_steps', 'close'],
  marketing_strategy: ['cover', 'context', 'audience', 'insights', 'strategy_pillars', 'big_idea', 'channel_plan', 'content_system', 'measurement', 'timeline', 'next_steps', 'close'],
  qbr: ['cover', 'agenda', 'kpi_dashboard', 'wins', 'losses', 'insights', 'pipeline', 'priorities_next_quarter', 'asks', 'close'],
  product_roadmap: ['cover', 'vision', 'now_next_later', 'timeline', 'themes', 'milestones', 'dependencies_risks', 'asks', 'close'],
  company_profile: ['cover', 'who_we_are', 'mission_vision', 'capabilities', 'proof_clients', 'case_study', 'team', 'process', 'cta', 'close'],
  training_workshop: ['cover', 'objectives', 'agenda', 'concepts', 'steps', 'exercise', 'quiz', 'summary', 'cta', 'close'],
  project_status_update: ['cover', 'agenda', 'traffic_light', 'progress', 'blockers', 'timeline', 'decisions', 'next_steps', 'close'],
  keynote_thought_leadership: ['cover', 'hook', 'tension', 'insight', 'big_idea', 'proof', 'implications', 'call_to_action', 'close']
};

// --- Strict structure for ad-agency creative decks (10 slides, fixed order) ---
// This is intentionally *rigid* because agency decks often need a reliable client-friendly narrative.
const AGENCY_CREATIVE_10 = [
  { id: 's1_title', name: 'Title', kind: 'title', slide_kinds: ['title', 'cover'] },
  { id: 's2_current_audience', name: 'Current Audience', kind: 'current_audience', slide_kinds: ['current_audience', 'audience'] },
  { id: 's3_about_brand', name: 'About the Brand', kind: 'about_brand', slide_kinds: ['about_brand', 'brand'] },
  { id: 's4_challenge', name: 'The Challenge / The Problem', kind: 'challenge', slide_kinds: ['challenge', 'problem'] },
  { id: 's5_opportunity', name: 'The Opportunity', kind: 'opportunity', slide_kinds: ['opportunity', 'insight_or_reframe'] },
  { id: 's6_pillars', name: 'Communication Pillars / Media Types', kind: 'communication_pillars', slide_kinds: ['communication_pillars', 'strategy_pillars'] },
  { id: 's7_concept', name: 'Creative Concept', kind: 'creative_concept', slide_kinds: ['creative_concept', 'big_idea'] },
  { id: 's8_visual_identity', name: 'Visual Identity', kind: 'visual_identity', slide_kinds: ['visual_identity', 'visual_direction'] },
  { id: 's9_execution', name: 'Execution Example', kind: 'execution_example', slide_kinds: ['execution_example', 'execution_examples'] },
  { id: 's10_thanks', name: 'Thank You', kind: 'thank_you', slide_kinds: ['thank_you', 'close'] }
];

function shouldLockAgency10(extractJson, options = {}) {
  const requested = asStr(options.deckType || options.deck_type || '', 80).trim();
  if (requested.toLowerCase() === 'ad_agency') return true;
  const suggested = asStr(extractJson?.deck_type_suggestion || extractJson?.deck_type || '', 80).trim();
  return suggested.toLowerCase() === 'ad_agency';
}

function lockNarrativeToAgency10(narrativeJson, extractJson) {
  const thesis = asStr(narrativeJson?.thesis || extractJson?.objective || extractJson?.title || 'A campaign story that earns attention.', 220);
  return {
    ...narrativeJson,
    deck_type: 'ad_agency',
    recipe_name: 'agency_creative_10',
    thesis,
    sections: AGENCY_CREATIVE_10.map((s, idx) => ({
      id: s.id,
      name: s.name,
      goal: `Deliver ${s.name.toLowerCase()} with one clear idea.`,
      key_message: '',
      must_include: [],
      slide_kinds: s.slide_kinds,
      transition_to_next: idx === AGENCY_CREATIVE_10.length - 1 ? '' : 'Next: we build the case.'
    }))
  };
}

function isAgencyTypographic(options = {}) {
  const ds = asStr(options.deckStyle || options.deck_style || '', 80).trim().toLowerCase();
  return ds === 'agency_typographic';
}

function lockDeckPlanToAgency10(deckPlan, extractJson, options = {}) {
  const title = asStr(deckPlan?.deck_title || extractJson?.title || 'Creative Campaign', 120);
  const subtitle = asStr(deckPlan?.deck_subtitle || extractJson?.subtitle || '', 140);

  const typographic = isAgencyTypographic(options);

  // Map desired slide kind -> default layout
  const defaultLayoutByKind = typographic
    ? {
      title: 'agency_center',
      cover: 'agency_center',
      current_audience: 'agency_infographic',
      audience: 'agency_infographic',
      about_brand: 'agency_half',
      brand: 'agency_half',
      challenge: 'agency_half',
      problem: 'agency_half',
      opportunity: 'agency_infographic',
      insight_or_reframe: 'agency_infographic',
      communication_pillars: 'agency_infographic',
      strategy_pillars: 'agency_infographic',
      creative_concept: 'agency_center',
      big_idea: 'agency_center',
      visual_identity: 'agency_half',
      visual_direction: 'agency_half',
      execution_example: 'full_bleed',
      execution_examples: 'full_bleed',
      thank_you: 'agency_center',
      close: 'agency_center'
    }
    : {
      title: 'hero',
      cover: 'hero',
      current_audience: 'cards',
      audience: 'cards',
      about_brand: 'split',
      brand: 'split',
      challenge: 'full_bleed',
      problem: 'full_bleed',
      opportunity: 'two_column',
      insight_or_reframe: 'two_column',
      communication_pillars: 'process_steps',
      strategy_pillars: 'process_steps',
      creative_concept: 'hero',
      big_idea: 'hero',
      visual_identity: 'image_caption',
      visual_direction: 'image_caption',
      execution_example: 'full_bleed',
      execution_examples: 'full_bleed',
      thank_you: 'hero',
      close: 'hero'
    };

  const defaultImagePromptByKind = {
    title: 'High-contrast abstract campaign key visual, premium editorial lighting, minimal, no text',
    current_audience: 'Modern editorial audience collage, diverse silhouettes, premium lighting, minimal, no text',
    about_brand: 'Premium brand essence key visual, clean minimal composition, no text',
    challenge: 'Dramatic abstract tension visual, high contrast, minimal, no text',
    opportunity: 'Optimistic breakthrough abstract visual, premium lighting, minimal, no text',
    communication_pillars: 'Minimal bento-style abstract icons and shapes, premium, high contrast, no text',
    creative_concept: 'Signature campaign platform key visual, iconic, bold, minimal, no text',
    visual_identity: 'Design system moodboard: materials, textures, color swatches, minimal, no text',
    execution_example: 'Cinematic outdoor advertising mockup scene, generic, premium lighting, no logos, no text',
    thank_you: 'Soft gradient background, premium minimal, no text'
  };

  const existing = Array.isArray(deckPlan?.slides) ? deckPlan.slides : [];
  const used = new Set();

  const takeFirstMatch = (kinds) => {
    for (let i = 0; i < existing.length; i++) {
      if (used.has(i)) continue;
      const k = (existing[i]?.kind || '').toString().toLowerCase();
      if (kinds.map(x=>x.toLowerCase()).includes(k)) {
        used.add(i);
        return existing[i];
      }
    }
    return null;
  };

  const slides = AGENCY_CREATIVE_10.map((s, idx) => {
    const match = takeFirstMatch(s.slide_kinds);
    const baseKind = s.kind;
    const layout = match?.layout || defaultLayoutByKind[baseKind] || 'split';

    // Create safe placeholders if missing (editor pass will polish messaging)
    const safe = {
      kind: baseKind,
      section: s.name,
      layout,
      title: match?.title || (baseKind === 'title' ? title : s.name),
      subtitle: match?.subtitle || (baseKind === 'title' ? subtitle : ''),
      bullets: Array.isArray(match?.bullets) ? match.bullets : [],
      stat: match?.stat ?? null,
      quote: match?.quote ?? null,
      image_prompt: typeof match?.image_prompt === 'string'
        ? match.image_prompt
        : (defaultImagePromptByKind[baseKind] || 'High-contrast abstract campaign key visual, premium editorial lighting, minimal, no text'),
      speaker_notes: match?.speaker_notes || '',
      setup_line: match?.setup_line || '',
      takeaway: match?.takeaway || '',
      bridge_line: match?.bridge_line || (idx === AGENCY_CREATIVE_10.length - 1 ? '' : 'Next:')
    };

    // Force creative concept slide behavior: short punchy line as title
    if (baseKind === 'creative_concept') {
      safe.subtitle = safe.subtitle || 'The big idea in one line.';
      safe.bullets = safe.bullets?.slice(0, 3);
    }

    // Agency typographic mode: keep slides sparse (big type, short subhead, minimal support lines)
    if (typographic) {
      const maxBullets = (baseKind === 'communication_pillars' || baseKind === 'current_audience' || baseKind === 'opportunity') ? 3 : 2;
      safe.bullets = Array.isArray(safe.bullets) ? safe.bullets.slice(0, maxBullets) : [];
    }
    return safe;
  });

  return {
    ...deckPlan,
    deck_type: 'ad_agency',
    deck_title: title,
    deck_subtitle: subtitle,
    recommended_slide_count: 10,
    slides
  };
}
/**
 * --- Agency concept quality pass ---
 * We keep this *very* scoped: only touches the Creative Concept slide messaging.
 * (No structural changes, no global bullet/length enforcement.)
 */
function isWeakAgencyConceptLine(line = '', briefText = '') {
  const s = (line || '').trim();
  if (!s) return true;

  // 1) Ban label-style concepts
  if (s.includes(':')) return true;

  // 2) Ban generic "Campaign:" / "Out Of Home Campaign" phrasing
  if (/\bcampaign\b/i.test(s) && /\b(out of home|ooh)\b/i.test(s)) return true;
  if (/^\s*\w+\s+campaign\b/i.test(s)) return true;

  // 3) Ban alliteration list concepts like "Visible, Vibrant, Victorious"
  const parts = s.split(',').map(p => p.trim()).filter(Boolean);
  if (parts.length >= 3) {
    const first = (parts[0][0] || '').toLowerCase();
    const allSame = first && parts.every(p => (p[0] || '').toLowerCase() === first);
    if (allSame) return true;
  }

  // 4) If it's too long, it's usually not a platform line
  const words = s.replace(/[.·•]/g, ' ').trim().split(/\s+/).filter(Boolean);
  if (words.length > 12) return true;

  // 5) If it has zero overlap with the brief's key terms, it's likely generic
  const kw = briefKeywords(briefText);
  if (kw.size) {
    const sNorm = s.toLowerCase();
    const hit = Array.from(kw).some(k => sNorm.includes(k));
    if (!hit && words.length <= 6) return true; // short + no overlap = generic
  }

  return false;
}

function briefKeywords(text = '') {
  const t = (text || '').toLowerCase();
  const whitelist = [
    'radio','frequency','frequencies','city','cities','tune','tuned','vibe','vibes','listen','listening','airwaves','station','stations',
    'audience','fans','music','beats','sound','saudi','riyadh','jeddah','dammam','khobar','ooh','out of home','campaign'
  ];
  const found = new Set();
  for (const w of whitelist) {
    if (t.includes(w)) found.add(w.split(' ')[0]); // keep single token
  }
  return found;
}

function buildAgencyConceptSchema() {
  return {
    name: 'agency_concept_refine',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        keep_existing: { type: 'boolean' },
        concept_line: { type: 'string' },
        supporting_line: { type: 'string' },
        proof_points: { type: 'array', items: { type: 'string' }, minItems: 3, maxItems: 4 },
        rationale: { type: 'string' }
      },
      required: ['keep_existing', 'concept_line', 'supporting_line', 'proof_points', 'rationale']
    }
  };
}

function buildAgencyConceptSystemPrompt({ language = 'English' }) {
  return [
    `You are an award-winning creative director. You write REAL campaign platform lines — not labels.`,
    `Task: refine ONLY the Creative Concept slide in an agency deck.`,
    `Output JSON must match the provided schema exactly.`,
    `Hard rules for concept_line:`,
    `- Must be a single campaign platform line (2–8 words preferred; up to 12 max).`,
    `- May use punctuation (e.g., "X. Y.") — encouraged if it helps memorability.`,
    `- NO colon ":" anywhere. NO "Campaign:" prefixes. NO generic alliteration lists (e.g., "Vivid, Vibrant, Victorious").`,
    `- Must be clearly on-brief: echo the core tension + promise from the brief (what we are trying to change + why it matters).`,
    `- Do not invent facts, numbers, results, partners, or dates.`,
    `If the current concept line is already excellent AND tightly on-brief, set keep_existing=true and repeat it verbatim.`,
    `Language: ${language}.`
  ].join(' ');
}

function buildAgencyConceptUserPrompt({ briefText, extractJson, narrativeJson, messagingMap, currentConceptSlide }) {
  const current = currentConceptSlide || {};
  return [
    `BRIEF (verbatim):`,
    briefText,
    `\n\nExtracted brief JSON:`,
    JSON.stringify(extractJson, null, 2),
    extractJson?.explicit_outline?.length ? `\nExplicit slide outline detected in brief (use it as backbone; expand naturally):\n${extractJson.explicit_outline.map(o=>`- Slide ${o.slide}: ${o.title}`).join('\n')}` : '',
    `\n\nNarrative plan JSON:`,
    JSON.stringify(narrativeJson, null, 2),
    `\n\nMessaging map JSON:`,
    JSON.stringify(messagingMap, null, 2),
    `\n\nCurrent Creative Concept slide (only this slide may be changed):`,
    JSON.stringify(current, null, 2),
    `\n\nDeliver:`,
    `1) concept_line (the campaign platform line)`,
    `2) supporting_line (1 sentence: what it means + how it solves the brief)`,
    `3) proof_points (3–4 bullets: concrete reasons, grounded in the brief)`,
    `4) rationale (short: why this is better than the current line)`
  ].join('\n');
}

async function refineAgencyConceptWithOpenAI(briefText, extractJson, narrativeJson, messagingMap, deckPlan, options = {}) {
  const slides = Array.isArray(deckPlan?.slides) ? deckPlan.slides : [];
  const idx = slides.findIndex(s => ['creative_concept','big_idea'].includes((s?.kind || '').toString().toLowerCase()));
  if (idx < 0) return deckPlan;

  const current = slides[idx] || {};
  const currentLine = (current.title || '').toString();
  const needsHelp = isWeakAgencyConceptLine(currentLine, briefText);

  const client = getOpenAIClient();
  const { text: model } = getModels();
  const schema = buildAgencyConceptSchema();
  const system = buildAgencyConceptSystemPrompt({ language: asStr(options.language || 'English', 80) });
  const user = buildAgencyConceptUserPrompt({ briefText, extractJson, narrativeJson, messagingMap, currentConceptSlide: current });

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.6
  });

  const content = completion.choices?.[0]?.message?.content;
  if (!content) return deckPlan;

  const out = JSON.parse(content);
  const conceptLine = (out.concept_line || '').toString().trim();
  const keep = !!out.keep_existing;

  const finalLine = (keep && !needsHelp) ? currentLine : conceptLine;
  const patchedLine = isWeakAgencyConceptLine(finalLine, briefText) ? (conceptLine || currentLine) : finalLine;

  const nextSlides = slides.map((s, i) => {
    if (i !== idx) return s;
    return {
      ...(s || {}),
      title: patchedLine || currentLine || 'Big idea',
      subtitle: (out.supporting_line || s.subtitle || '').toString(),
      bullets: Array.isArray(out.proof_points) ? out.proof_points.filter(Boolean) : (Array.isArray(s.bullets) ? s.bullets : [])
    };
  });

  return { ...(deckPlan || {}), slides: nextSlides };
}

async function refineAgencyConceptWithGemini(briefText, extractJson, narrativeJson, messagingMap, deckPlan, options = {}) {
  const slides = Array.isArray(deckPlan?.slides) ? deckPlan.slides : [];
  const idx = slides.findIndex(s => ['creative_concept','big_idea'].includes((s?.kind || '').toString().toLowerCase()));
  if (idx < 0) return deckPlan;

  const current = slides[idx] || {};
  const currentLine = (current.title || '').toString();
  const needsHelp = isWeakAgencyConceptLine(currentLine, briefText);

  const { gemini_text: model } = getModels();
  const schema = buildAgencyConceptSchema();

  const system = buildAgencyConceptSystemPrompt({ language: asStr(options.language || 'English', 80) });
  const user = buildAgencyConceptUserPrompt({ briefText, extractJson, narrativeJson, messagingMap, currentConceptSlide: current });

  const out = await geminiGenerateJson({
    model,
    system,
    user,
    jsonSchema: schema.schema,
    temperature: 0.6
  });

  const conceptLine = (out?.concept_line || '').toString().trim();
  const keep = !!out?.keep_existing;

  const finalLine = (keep && !needsHelp) ? currentLine : conceptLine;
  const patchedLine = isWeakAgencyConceptLine(finalLine, briefText) ? (conceptLine || currentLine) : finalLine;

  const nextSlides = slides.map((s, i) => {
    if (i !== idx) return s;
    return {
      ...(s || {}),
      title: patchedLine || currentLine || 'Big idea',
      subtitle: (out?.supporting_line || s.subtitle || '').toString(),
      bullets: Array.isArray(out?.proof_points) ? out.proof_points.filter(Boolean) : (Array.isArray(s.bullets) ? s.bullets : [])
    };
  });

  return { ...(deckPlan || {}), slides: nextSlides };
}

function clampInt(n, lo, hi, fallback) {
  const x = Number.isFinite(+n) ? Math.floor(+n) : fallback;
  return Math.max(lo, Math.min(hi, x));
}


function extractExplicitOutlineSignals(briefText = '') {
  const text = (briefText || '').toString();
  const lines = text.split(/\r?\n/);
  const items = [];
  // Slide markers like: "Slide 1:", "Slide 1 -", "Page 3 —", "SLIDE 4 •"
  const re1 = /^\s*(slide|page)\s*(\d{1,3})\s*[:\-–•]\s*(.+?)\s*$/i;
  // Numbered list markers like: "1) Title", "2. Agenda"
  const re2 = /^\s*(\d{1,3})\s*[\).:-]\s*(.+?)\s*$/;
  for (const ln of lines) {
    const m1 = ln.match(re1);
    if (m1) {
      const n = parseInt(m1[2], 10);
      const title = (m1[3] || '').trim();
      if (Number.isFinite(n) && title) items.push({ n, title });
      continue;
    }
    const m2 = ln.match(re2);
    if (m2) {
      const n = parseInt(m2[1], 10);
      const title = (m2[2] || '').trim();
      // Avoid capturing random numbered paragraphs by requiring title-ish length
      if (Number.isFinite(n) && title && title.length <= 140) items.push({ n, title });
    }
  }

  // Deduplicate by slide number, keep first
  const byN = new Map();
  for (const it of items) if (!byN.has(it.n)) byN.set(it.n, it.title);
  const outline = Array.from(byN.entries()).sort((a,b)=>a[0]-b[0]).map(([n,title])=>({ slide: n, title }));

  // Detect range hints like "approx. 18–22 slides" or "18-22 slides"
  const rangeMatch = text.match(/(\d{1,2})\s*[–-]\s*(\d{1,2})\s*slides?/i);
  const range = rangeMatch ? { min: parseInt(rangeMatch[1],10), max: parseInt(rangeMatch[2],10) } : null;

  return { outline, range };
}

function asStr(v, max = 200) {
  return (v ?? '').toString().slice(0, max);
}

function resolveVoiceProfile(options = {}) {
  const raw = asStr(options.voiceProfile || options.voice_profile || options.voice || 'witty_agency', 40)
    .toLowerCase()
    .trim();
  return VOICE_PROFILES.includes(raw) ? raw : 'witty_agency';
}

function buildExtractSchema() {
  return {
    name: 'brief_extract',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        deck_type_suggestion: { type: 'string', enum: DECK_TYPES },
        title: { type: 'string' },
        subtitle: { type: 'string' },
        audience: { type: 'string' },
        language: { type: 'string' },
        vibe: { type: 'string' },

        // Core intent
        objective: { type: 'string' },
        ask_or_cta: { type: 'string' },
        constraints: { type: 'array', items: { type: 'string' }, maxItems: 12 },

        // Common sections (optional / best-effort)
        problem: { type: 'string' },
        solution: { type: 'string' },
        product_or_service: { type: 'string' },
        differentiators: { type: 'array', items: { type: 'string' }, maxItems: 10 },
        market: { type: 'string' },
        business_model: { type: 'string' },
        traction_or_proof: { type: 'array', items: { type: 'string' }, maxItems: 10 },
        competition: { type: 'array', items: { type: 'string' }, maxItems: 8 },

        // Ops / delivery
        scope: { type: 'array', items: { type: 'string' }, maxItems: 12 },
        deliverables: { type: 'array', items: { type: 'string' }, maxItems: 12 },
        timeline: {
          type: 'array',
          maxItems: 12,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              date_or_phase: { type: 'string' },
              label: { type: 'string' },
              detail: { type: 'string' }
            },
            required: ['date_or_phase', 'label', 'detail']
          }
        },
        risks: { type: 'array', items: { type: 'string' }, maxItems: 10 },

        // Metrics / QBR
        kpis: {
          type: 'array',
          maxItems: 10,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              label: { type: 'string' },
              value: { type: 'string' },
              delta: { type: 'string' }
            },
            required: ['label', 'value', 'delta']
          }
        },

        // Status updates
        status_items: {
          type: 'array',
          maxItems: 12,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              item: { type: 'string' },
              status: { type: 'string', enum: ['red', 'yellow', 'green'] },
              owner: { type: 'string' },
              eta: { type: 'string' },
              blocker: { type: 'string' }
            },
            required: ['item', 'status', 'owner', 'eta', 'blocker']
          }
        },

        // Commercials
        pricing: {
          type: ['object', 'null'],
          additionalProperties: false,
          properties: {
            currency: { type: 'string' },
            plans: {
              type: 'array',
              maxItems: 4,
              items: {
                type: 'object',
                additionalProperties: false,
                properties: {
                  name: { type: 'string' },
                  price: { type: 'string' },
                  period: { type: 'string' },
                  bullets: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  highlight: { type: 'boolean' }
                },
                required: ['name', 'price', 'period', 'bullets', 'highlight']
              }
            },
            notes: { type: 'string' }
          },
          required: ['currency', 'plans', 'notes']
        },

        // People
        team: {
          type: 'array',
          maxItems: 10,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              name: { type: 'string' },
              role: { type: 'string' },
              bio: { type: 'string' }
            },
            required: ['name', 'role', 'bio']
          }
        },

        // Missing info
        missing_info: { type: 'array', items: { type: 'string' }, maxItems: 12 },
        source_summary: { type: 'string' }
      },
      required: [
        'deck_type_suggestion',
        'title',
        'subtitle',
        'audience',
        'language',
        'vibe',
        'objective',
        'ask_or_cta',
        'constraints',
        'problem',
        'solution',
        'product_or_service',
        'differentiators',
        'market',
        'business_model',
        'traction_or_proof',
        'competition',
        'scope',
        'deliverables',
        'timeline',
        'risks',
        'kpis',
        'status_items',
        'pricing',
        'team',
        'missing_info',
        'source_summary'
      ]
    }
  };
}

function buildNarrativeSchema() {
  return {
    name: 'narrative_plan',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        deck_type: { type: 'string', enum: DECK_TYPES },
        recipe_name: { type: 'string' },
        thesis: { type: 'string' },
        story_spine: {
          type: 'object',
          additionalProperties: false,
          properties: {
            setup: { type: 'string' },
            tension: { type: 'string' },
            insight: { type: 'string' },
            solution: { type: 'string' },
            proof: { type: 'string' },
            action: { type: 'string' }
          },
          required: ['setup', 'tension', 'insight', 'solution', 'proof', 'action']
        },
        lexicon: {
          type: 'object',
          additionalProperties: false,
          properties: {
            prefer_terms: { type: 'array', items: { type: 'string' }, maxItems: 15 },
            avoid_terms: { type: 'array', items: { type: 'string' }, maxItems: 15 }
          },
          required: ['prefer_terms', 'avoid_terms']
        },
        sections: {
          type: 'array',
          minItems: 4,
          maxItems: 10,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              id: { type: 'string' },
              name: { type: 'string' },
              goal: { type: 'string' },
              key_message: { type: 'string' },
              must_include: { type: 'array', items: { type: 'string' }, maxItems: 6 },
              slide_kinds: { type: 'array', items: { type: 'string' }, maxItems: 8 },
              transition_to_next: { type: 'string' }
            },
            required: ['id', 'name', 'goal', 'key_message', 'must_include', 'slide_kinds', 'transition_to_next']
          }
        }
      },
      required: ['deck_type', 'recipe_name', 'thesis', 'story_spine', 'lexicon', 'sections']
    }
  };
}

function buildMessagingSchema() {
  return {
    name: 'messaging_map',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        voice_profile: { type: 'string', enum: VOICE_PROFILES },
        anchors: {
          type: 'array',
          minItems: 1,
          maxItems: 3,
          items: { type: 'string' }
        },
        locked_phrases: {
          type: 'array',
          minItems: 5,
          maxItems: 12,
          items: { type: 'string' }
        },
        section_concepts: {
          type: 'array',
          minItems: 4,
          maxItems: 10,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              section_id: { type: 'string' },
              concept: { type: 'string', description: 'The ONE new idea this section introduces (short claim).' },
              proof_points: { type: 'array', items: { type: 'string' }, maxItems: 4 },
              forbidden: { type: 'array', items: { type: 'string' }, maxItems: 6 },
              required_bridge: { type: 'string', description: 'A bridge line that cleanly tees up the next section.' }
            },
            required: ['section_id', 'concept', 'proof_points', 'forbidden', 'required_bridge']
          }
        },
        buzzwords_to_avoid: { type: 'array', items: { type: 'string' }, maxItems: 20 }
      },
      required: ['voice_profile', 'anchors', 'locked_phrases', 'section_concepts', 'buzzwords_to_avoid']
    }
  };
}

function buildDeckSchema() {
  return {
    name: 'deck_plan',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        deck_type: {
          type: 'string',
          enum: DECK_TYPES,
          description: 'Overall presentation type / intent.'
        },
        deck_title: { type: 'string' },
        deck_subtitle: { type: 'string' },
        recommended_slide_count: { type: 'integer', minimum: 5, maximum: SAFE_MAX_SLIDES },
        theme: {
          type: 'object',
          additionalProperties: false,
          properties: {
            vibe: { type: 'string' },
            primary_color: { type: 'string' },
            secondary_color: { type: 'string' },
            font_heading: { type: 'string' },
            font_body: { type: 'string' }
          },
          required: ['vibe', 'primary_color', 'secondary_color', 'font_heading', 'font_body']
        },
        slides: {
          type: 'array',
          minItems: 5,
          maxItems: SAFE_MAX_SLIDES,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              kind: { type: 'string', description: 'Semantic role (cover, agenda, problem, solution, KPI dashboard, timeline, pricing, CTA, etc.).' },
              layout: { type: 'string', enum: SLIDE_LAYOUTS },
              // Narrative glue (optional but used to enforce flow)
              section: { type: 'string', description: 'Which narrative section this slide belongs to (e.g., Challenge, Opportunity, Strategy, Creative Platform).' },
              setup_line: { type: 'string', description: '1 short sentence that links from the previous slide.' },
              takeaway: { type: 'string', description: 'The one sentence this slide should leave in the reader\'s head.' },
              bridge_line: { type: 'string', description: '1 short sentence that tees up the next slide.' },
              title: { type: 'string' },
              subtitle: { type: 'string' },
              bullets: { type: 'array', items: { type: 'string' }, maxItems: 8 },

              // Existing blocks
              stat: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: { value: { type: 'string' }, label: { type: 'string' } },
                required: ['value', 'label']
              },
              quote: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: { text: { type: 'string' }, attribution: { type: 'string' } },
                required: ['text', 'attribution']
              },

              // New: agenda
              agenda_items: { type: ['array', 'null'], items: { type: 'string' }, maxItems: 10 },

              // New: cards (features, personas, benefits)
              cards: {
                type: ['array', 'null'],
                maxItems: 6,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    title: { type: 'string' },
                    body: { type: 'string' },
                    tag: { type: 'string' }
                  },
                  required: ['title', 'body', 'tag']
                }
              },

              // New: timeline
              timeline_items: {
                type: ['array', 'null'],
                maxItems: 12,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    date_or_phase: { type: 'string' },
                    label: { type: 'string' },
                    detail: { type: 'string' }
                  },
                  required: ['date_or_phase', 'label', 'detail']
                }
              },

              // New: KPI dashboard tiles
              kpis: {
                type: ['array', 'null'],
                maxItems: 8,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    label: { type: 'string' },
                    value: { type: 'string' },
                    delta: { type: 'string' }
                  },
                  required: ['label', 'value', 'delta']
                }
              },

              // New: traffic light status table
              status_items: {
                type: ['array', 'null'],
                maxItems: 12,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    item: { type: 'string' },
                    status: { type: 'string', enum: ['red', 'yellow', 'green'] },
                    owner: { type: 'string' },
                    eta: { type: 'string' },
                    blocker: { type: 'string' }
                  },
                  required: ['item', 'status', 'owner', 'eta', 'blocker']
                }
              },

              // New: generic table
              table: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  headers: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 6 },
                  rows: {
                    type: 'array',
                    maxItems: 12,
                    items: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 6 }
                  }
                },
                required: ['headers', 'rows']
              },

              // New: pricing plans
              pricing: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  currency: { type: 'string' },
                  plans: {
                    type: 'array',
                    maxItems: 4,
                    items: {
                      type: 'object',
                      additionalProperties: false,
                      properties: {
                        name: { type: 'string' },
                        price: { type: 'string' },
                        period: { type: 'string' },
                        bullets: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                        highlight: { type: 'boolean' }
                      },
                      required: ['name', 'price', 'period', 'bullets', 'highlight']
                    }
                  },
                  notes: { type: 'string' }
                },
                required: ['currency', 'plans', 'notes']
              },

              // New: comparison matrix
              matrix: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  x_labels: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 6 },
                  y_labels: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 8 },
                  cells: {
                    type: 'array',
                    maxItems: 8,
                    items: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 6 }
                  }
                },
                required: ['x_labels', 'y_labels', 'cells']
              },

              // New: steps / process
              steps: {
                type: ['array', 'null'],
                maxItems: 8,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    title: { type: 'string' },
                    detail: { type: 'string' }
                  },
                  required: ['title', 'detail']
                }
              },

              // New: people (team slide)
              people: {
                type: ['array', 'null'],
                maxItems: 10,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    name: { type: 'string' },
                    role: { type: 'string' },
                    bio: { type: 'string' }
                  },
                  required: ['name', 'role', 'bio']
                }
              },

              // New: logo wall (text-only)
              logo_items: { type: ['array', 'null'], items: { type: 'string' }, maxItems: 30 },

              // New: explicit CTA
              cta: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  headline: { type: 'string' },
                  primary_action: { type: 'string' },
                  secondary_action: { type: 'string' }
                },
                required: ['headline', 'primary_action', 'secondary_action']
              },

              // New: SWOT
              swot: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  strengths: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  weaknesses: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  opportunities: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  threats: { type: 'array', items: { type: 'string' }, maxItems: 8 }
                },
                required: ['strengths', 'weaknesses', 'opportunities', 'threats']
              },

              // New: Funnel
              funnel: {
                type: ['array', 'null'],
                maxItems: 7,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: { label: { type: 'string' }, value: { type: 'string' } },
                  required: ['label', 'value']
                }
              },

              // New: Now / Next / Later
              now_next_later: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  now: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  next: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  later: { type: 'array', items: { type: 'string' }, maxItems: 8 }
                },
                required: ['now', 'next', 'later']
              },

              // New: OKRs
              okrs: {
                type: ['array', 'null'],
                maxItems: 5,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    objective: { type: 'string' },
                    key_results: { type: 'array', items: { type: 'string' }, maxItems: 6 }
                  },
                  required: ['objective', 'key_results']
                }
              },

              // New: Case study
              case_study: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  client: { type: 'string' },
                  challenge: { type: 'string' },
                  approach: { type: 'array', items: { type: 'string' }, maxItems: 8 },
                  results: { type: 'array', items: { type: 'string' }, maxItems: 8 }
                },
                required: ['client', 'challenge', 'approach', 'results']
              },

              // New: Mermaid diagram
              diagram: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  code: { type: 'string' },
                  theme: { type: 'string' }
                },
                required: ['code', 'theme']
              },

              // New: Icon list (Iconify names)
              icons: {
                type: ['array', 'null'],
                maxItems: 6,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: {
                    name: { type: 'string' },
                    label: { type: 'string' }
                  },
                  required: ['name', 'label']
                }
              },

              // New: Simple charts
              chart: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  chart_type: { type: 'string', enum: ['bar', 'line'] },
                  spec: { type: ['object', 'null'] },
                  labels: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 10 },
                  values: { type: 'array', items: { type: 'number' }, minItems: 2, maxItems: 10 },
                  value_suffix: { type: 'string' }
                },
                required: ['chart_type', 'labels', 'values', 'value_suffix']
              },

              // New: Org chart (simple)
              org_chart: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: {
                  head: { type: 'string' },
                  reports: { type: 'array', items: { type: 'string' }, minItems: 2, maxItems: 8 }
                },
                required: ['head', 'reports']
              },

              // New: FAQ
              faq: {
                type: ['array', 'null'],
                maxItems: 8,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: { q: { type: 'string' }, a: { type: 'string' } },
                  required: ['q', 'a']
                }
              },

              image_prompt: {
                type: 'string',
                description: 'Prompt to generate a supporting background/illustration. Use empty string or "NONE" if no image is needed.'
              },
              speaker_notes: { type: 'string' }
            },
            required: ['kind', 'layout', 'section', 'setup_line', 'takeaway', 'bridge_line', 'title', 'subtitle', 'bullets', 'stat', 'quote', 'agenda_items', 'cards', 'timeline_items', 'kpis', 'status_items', 'table', 'pricing', 'matrix', 'steps', 'people', 'logo_items', 'cta', 'swot', 'funnel', 'now_next_later', 'okrs', 'case_study', 'chart', 'org_chart', 'faq', 'image_prompt', 'speaker_notes']
          }
        }
      },
      required: ['deck_type', 'deck_title', 'deck_subtitle', 'recommended_slide_count', 'theme', 'slides']
    }
  };
}

function buildExtractSystemPrompt({ vibe, audience, language }) {
  return [
    `You are a senior strategist. Extract structured facts from a messy brief so another model can build a deck.`,
    `Be faithful to the brief. Do NOT invent specific numbers, dates, or claims. Use placeholders if missing.`,
    `Output MUST match the JSON schema strictly.`,
    `Keep text short and usable.`,
    `Language: ${language}. Audience: ${audience}. Desired vibe: ${vibe}.`
  ].join(' ');
}

function buildNarrativeSystemPrompt({ vibe, audience, language, voiceProfile, requestedDeckType }) {
  // Keep this prompt deterministic: it produces the *journey blueprint*.
  const recipeNames = Object.keys(DECK_RECIPES);
  const recipeLines = recipeNames
    .map((k) => `${k}: ${DECK_RECIPES[k].join(' → ')}`)
    .join(' | ');
  const voice = VOICE_RULES[voiceProfile] || VOICE_RULES.witty_agency;
  return [
    `You are a senior strategist and narrative director. Your job is to create the messaging journey for a deck.`,
    `Input: extracted brief JSON (facts). Output: a narrative_plan that ensures slide-to-slide flow (no random jumps).`,
    `Do NOT invent hard facts. If something is unknown, keep it generic and add it to must_include as a question/placeholder.`,
    `Choose a recipe_name from: ${recipeNames.join(', ')}.`,
    `Recipe sequences: ${recipeLines}`,
    `Recipes are story templates (order of ideas), not visual templates.`,
    `IMPORTANT: If the deck_type is "ad_agency" (or the user requested ad_agency), you MUST use recipe_name "agency_creative_10" and create EXACTLY 10 sections matching these beats in this exact order: 1) Title 2) Current Audience 3) About the Brand 4) The Challenge/The Problem 5) The Opportunity 6) Communication Pillars/Media Types 7) Creative Concept (one punchy campaign line / big idea) 8) Visual Identity (proposed design style) 9) Execution Example 10) Thank You.`,
    `For marketing case studies, follow this proven arc: proof of signal → brand alignment → challenge → reframe to opportunity → pillars → platform → visual system → executions → measurement → next steps → close.`,
    `Write crisp, declarative key_message lines. Each section must end with a transition_to_next that tees up the next beat.`,
    `Voice profile: ${voice.name}. ${voice.tagline}`,
    `Headline discipline: ${voice.headline_rules.join(' ')} `,
    requestedDeckType ? `Requested deck type (if any): ${requestedDeckType}.` : `No explicit deck type requested.`,
    `Language: ${language}. Audience: ${audience}. Vibe: ${vibe}.`
  ].join(' ');
}

function buildNarrativeUserPrompt(extractJson) {
  return [
    `Extracted brief JSON (source of truth):`,
    JSON.stringify(extractJson, null, 2),
    extractJson?.explicit_outline?.length ? `\nExplicit slide outline detected in brief (use it as backbone; expand naturally):\n${extractJson.explicit_outline.map(o=>`- Slide ${o.slide}: ${o.title}`).join('\n')}` : '',
    `\nNow output narrative_plan JSON per schema.`
  ].join('\n\n');
}

function buildAssembleSystemPrompt({ vibe, audience, language, requestedDeckType, requestedSlides, voiceProfile }) {
  const voice = VOICE_RULES[voiceProfile] || VOICE_RULES.witty_agency;
  return [
    `You are a senior creative director and presentation architect.`,
    `Task: assemble a coherent slide-by-slide PowerPoint plan from: (1) extracted brief JSON (facts) and (2) narrative_plan (journey).`,
    `Output MUST follow the JSON schema strictly.`,
    `CRITICAL: Follow narrative_plan.sections in order. Do not shuffle beats. Do not introduce new topics late.`,
    `CRITICAL: Use narrative_plan.lexicon.prefer_terms and avoid narrative_plan.lexicon.avoid_terms to keep terminology consistent.`,
    `CRITICAL: Fill slide.section, slide.setup_line, slide.takeaway, slide.bridge_line to create smooth continuity.`,
    requestedDeckType === 'ad_agency'
      ? `AGENCY MODE (STRICT): Output EXACTLY 10 slides, in this exact order, with these slide.kind values: title, current_audience, about_brand, challenge, opportunity, communication_pillars, creative_concept, visual_identity, execution_example, thank_you. Do NOT add or remove slides. Do NOT rename kinds. Each slide.section should match the human-friendly page name.`
      : ``,
    `Prefer the new business layouts when appropriate: timeline, kpi_dashboard, traffic_light, table, pricing, comparison_matrix, process_steps, team_grid, logo_wall, agenda, section_header, swot, funnel, now_next_later, okr, case_study, chart_bar, chart_line, org_chart, faq, infographic_3.`,
    `Avoid filler. Make it crisp and executive-ready.`,
    `Bullets must carry meaning: each bullet should be 8–18 words, specific, and add new information (no stubs like "Increase awareness").`,
    `Never output empty bullets. No placeholder bullets like "TBD" inside bullets — use speaker_notes for unknowns.`,
    `If layout is "two_column": output EXACTLY 8 bullets where bullets[0..3] are KEY POINTS (8–12 words each) and bullets[4..7] are MORE DETAIL (12–20 words each with a concrete example, mechanism, or implication).`,
    `For list-heavy layouts (process_steps, timeline, cards, kpi_dashboard, traffic_light), every item must include an insight sentence, not just a label.`,
    `Use slide.kind for semantics (cover, agenda, problem, solution, KPI dashboard, roadmap, status, pricing, CTA, etc.).`,
    `Write headlines as declarative claims that advance the story (not labels). Each slide must support its section's key_message.`,
    `Use repetition with intent: introduce a core phrase/platform once, then echo it in later slides as a callback (without becoming repetitive).`,
    `Voice profile: ${voice.name}. ${voice.tagline}`,
    `Headline rules: ${voice.headline_rules.join(' ')} `,
    `Subhead rules: ${voice.subhead_rules.join(' ')} `,
    `Section rules: ${voice.section_rules.join(' ')} `,
    `Diction rules: ${voice.diction_rules.join(' ')} `,
    `Make slide variety: alternate layouts, include at least one data-style slide when relevant.`,
    `If data is missing, use safe placeholders and explain missing inputs in speaker_notes.`,
    requestedDeckType ? `Requested deck type: ${requestedDeckType} (respect unless clearly wrong).` : `Deck type: infer from extracted brief.`,
    requestedSlides ? `Target slides: ${requestedSlides} (soft target; keep structure coherent).` : `Choose 5–18 slides as needed.`,
    `Language: ${language}. Audience: ${audience}. Vibe: ${vibe}.`,
    `Image prompts: visually specific, no logos, no copyrighted characters, NEVER ask for text/words in images. For data/table slides, you may set image_prompt to "NONE".`
  ].join(' ');
}

function buildUserPrompt(briefText) {
  return `Brief (may include messy notes; extract intent, facts, and structure):\n\n${briefText}`;
}

function buildAssembleUserPrompt(extractJson, narrativeJson, messagingMap) {
  return [
    `Here is the extracted brief JSON (source of truth):`,
    JSON.stringify(extractJson, null, 2),
    extractJson?.explicit_outline?.length ? `\nExplicit slide outline detected in brief (use it as backbone; expand naturally):\n${extractJson.explicit_outline.map(o=>`- Slide ${o.slide}: ${o.title}`).join('\n')}` : '',
    `\nHere is the narrative_plan JSON (journey blueprint; follow section order):`,
    JSON.stringify(narrativeJson, null, 2),
    `\nHere is the messaging_map JSON (copy constraints; anchors + section concepts; you MUST comply):`,
    JSON.stringify(messagingMap, null, 2)
  ].join('\n\n');
}

function buildMessagingSystemPrompt({ vibe, audience, language, voiceProfile }) {
  const voice = VOICE_RULES[voiceProfile] || VOICE_RULES.witty_agency;
  return [
    `You are a creative director and narrative copy editor. Your job is to define the "agency brain" messaging constraints for this deck so the copy feels authored (not random).`,
    `Input: extracted brief JSON and narrative_plan JSON. Output: messaging_map JSON per schema.`,
    `Do NOT invent hard facts. Anchors can be crafted as brand/platform phrases, but must be consistent with the brief and narrative.`,
    `CRITICAL: For each narrative section, output EXACTLY ONE concept (a short claim) that section introduces. Everything else is proof or example.`,
    `CRITICAL: Provide a required_bridge for every section that tees up the next section cleanly (witty but not try-hard).`,
    `Voice profile: ${voice.name}. ${voice.tagline}`,
    `Headline rules: ${voice.headline_rules.join(' ')} `,
    `Section rules: ${voice.section_rules.join(' ')} `,
    `Diction rules: ${voice.diction_rules.join(' ')} `,
    `Buzzwords to avoid: ${(voice.forbidden_terms || []).join(', ') || 'None'}.`,
    `Language: ${language}. Audience: ${audience}. Vibe: ${vibe}.`
  ].join(' ');
}

function buildMessagingUserPrompt(extractJson, narrativeJson, voiceProfile) {
  return [
    `Extracted brief JSON (facts):`,
    JSON.stringify(extractJson, null, 2),
    extractJson?.explicit_outline?.length ? `\nExplicit slide outline detected in brief (use it as backbone; expand naturally):\n${extractJson.explicit_outline.map(o=>`- Slide ${o.slide}: ${o.title}`).join('\n')}` : '',
    `\nNarrative_plan JSON (journey blueprint):`,
    JSON.stringify(narrativeJson, null, 2),
    `\nRequested voice_profile: ${voiceProfile}`,
    `\nNow output messaging_map JSON per schema.`
  ].join('\n\n');
}

function buildEditSystemPrompt({ vibe, audience, language, voiceProfile, requestedDeckType }) {
  const voice = VOICE_RULES[voiceProfile] || VOICE_RULES.witty_agency;
  return [
    `You are the final creative director pass ("deck editor"). Your job is to rewrite ONLY the messaging so the deck reads as one authored journey.`,
    `Input: extracted brief JSON, narrative_plan JSON, messaging_map JSON, and an initial deck_plan JSON. Output: an improved deck_plan JSON matching the SAME schema exactly.`,
    requestedDeckType === 'ad_agency'
      ? `AGENCY MODE (STRICT): Keep the deck at EXACTLY 10 slides. Do NOT merge, split, add, or remove slides. Only rewrite the messaging.`
      : `DO NOT change the deck structure unless needed for flow: you may merge or split slides ONLY if absolutely necessary; prefer rewriting instead. Keep slide count within 5–18.`,
    `DO NOT invent new facts, numbers, dates, client names, results. If unknown, keep placeholders in speaker_notes.`,
    `CRITICAL RULES (must pass):`,
    `1) Headline = declarative claim (no labels).`,
    `2) Each section introduces one concept only (use messaging_map.section_concepts).`,
    `3) Bridges: every section ends with a bridge_line that tees up the next beat. Use messaging_map.required_bridge verbatim where possible.`,
    `4) Anchors: repeat messaging_map.anchors verbatim as callbacks (no paraphrases).`,
    `5) Consistent lexicon: prefer narrative_plan.lexicon.prefer_terms; avoid narrative_plan.lexicon.avoid_terms + messaging_map.buzzwords_to_avoid.`,
    `6) Bullet density: rewrite any bullet that is shorter than 6 words or feels generic. Every bullet should be specific, insightful, and self-contained.`,
    `7) Two-column slides: ensure the first half are KEY POINTS and the second half are MORE DETAIL elaborations. No single-word bullets.`,
    `Voice profile: ${voice.name}. ${voice.tagline}`,
    `Diction rules: ${voice.diction_rules.join(' ')} `,
    `Language: ${language}. Audience: ${audience}. Vibe: ${vibe}.`
  ].join(' ');
}

function buildEditUserPrompt(extractJson, narrativeJson, messagingMap, deckPlan) {
  return [
    `Extracted brief JSON:`,
    JSON.stringify(extractJson, null, 2),
    extractJson?.explicit_outline?.length ? `\nExplicit slide outline detected in brief (use it as backbone; expand naturally):\n${extractJson.explicit_outline.map(o=>`- Slide ${o.slide}: ${o.title}`).join('\n')}` : '',
    `\nNarrative_plan JSON:`,
    JSON.stringify(narrativeJson, null, 2),
    `\nMessaging_map JSON:`,
    JSON.stringify(messagingMap, null, 2),
    `\nInitial deck_plan JSON to edit:`,
    JSON.stringify(deckPlan, null, 2)
  ].join('\n\n');
}

async function extractBriefWithOpenAI(briefText, options = {}) {
  const client = getOpenAIClient();
  const { text: model } = getModels();

  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);

  const schema = buildExtractSchema();
  const system = buildExtractSystemPrompt({ vibe, audience, language });
  const user = buildUserPrompt(briefText);

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.4
  });

  const content = completion.choices?.[0]?.message?.content;
  if (!content) throw new Error('No extract returned from model.');
  return JSON.parse(content);
}

async function extractBriefWithGemini(briefText, options = {}) {
  const { gemini_text: defaultGeminiTextModel } = getModels();

  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);

  const schema = buildExtractSchema();
  const system = buildExtractSystemPrompt({ vibe, audience, language });
  const user = buildUserPrompt(briefText);

  return geminiGenerateJson({
    model: defaultGeminiTextModel,
    system,
    user,
    jsonSchema: schema.schema,
    temperature: 0.4
  });
}

async function planNarrativeWithOpenAI(extractJson, options = {}) {
  const client = getOpenAIClient();
  const { text: model } = getModels();

  const voiceProfile = resolveVoiceProfile(options);

  const vibe = asStr(options.vibe || extractJson?.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || extractJson?.audience || 'general', 120);
  const language = asStr(options.language || extractJson?.language || 'English', 80);

  const schema = buildNarrativeSchema();
  const system = buildNarrativeSystemPrompt({
    vibe,
    audience,
    language,
    voiceProfile,
    requestedDeckType: asStr(options.deckType || options.deck_type || '', 80).trim() || null
  });
  const user = buildNarrativeUserPrompt(extractJson);

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.35
  });

  const content = completion.choices?.[0]?.message?.content;
  if (!content) throw new Error('No narrative_plan returned from model.');
  return JSON.parse(content);
}

async function planNarrativeWithGemini(extractJson, options = {}) {
  const { gemini_text: defaultGeminiTextModel } = getModels();

  const voiceProfile = resolveVoiceProfile(options);

  const vibe = asStr(options.vibe || extractJson?.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || extractJson?.audience || 'general', 120);
  const language = asStr(options.language || extractJson?.language || 'English', 80);

  const schema = buildNarrativeSchema();
  const system = buildNarrativeSystemPrompt({
    vibe,
    audience,
    language,
    voiceProfile,
    requestedDeckType: asStr(options.deckType || options.deck_type || '', 80).trim() || null
  });
  const user = buildNarrativeUserPrompt(extractJson);

  return geminiGenerateJson({
    model: defaultGeminiTextModel,
    system,
    user,
    jsonSchema: schema.schema,
    temperature: 0.35
  });
}

async function planMessagingWithOpenAI(extractJson, narrativeJson, options = {}) {
  const client = getOpenAIClient();
  const { text: model } = getModels();

  const voiceProfile = resolveVoiceProfile(options);
  const vibe = asStr(options.vibe || extractJson?.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || extractJson?.audience || 'general', 120);
  const language = asStr(options.language || extractJson?.language || 'English', 80);

  const schema = buildMessagingSchema();
  const system = buildMessagingSystemPrompt({ vibe, audience, language, voiceProfile });
  const user = buildMessagingUserPrompt(extractJson, narrativeJson, voiceProfile);

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.35
  });

  const content = completion.choices?.[0]?.message?.content;
  if (!content) throw new Error('No messaging_map returned from model.');
  return JSON.parse(content);
}

async function planMessagingWithGemini(extractJson, narrativeJson, options = {}) {
  const { gemini_text: defaultGeminiTextModel } = getModels();

  const voiceProfile = resolveVoiceProfile(options);
  const vibe = asStr(options.vibe || extractJson?.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || extractJson?.audience || 'general', 120);
  const language = asStr(options.language || extractJson?.language || 'English', 80);

  const schema = buildMessagingSchema();
  const system = buildMessagingSystemPrompt({ vibe, audience, language, voiceProfile });
  const user = buildMessagingUserPrompt(extractJson, narrativeJson, voiceProfile);

  return geminiGenerateJson({
    model: defaultGeminiTextModel,
    system,
    user,
    jsonSchema: schema.schema,
    temperature: 0.35
  });
}

async function assembleDeckWithOpenAI(extractJson, narrativeJson, messagingMap, options = {}) {
  const client = getOpenAIClient();
  const { text: model } = getModels();

  const nSlides = clampInt(options.nSlides ?? 10, 5, SAFE_MAX_SLIDES, 10);
  const voiceProfile = resolveVoiceProfile(options);
  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);
  const requestedDeckTypeRaw = asStr(options.deckType || options.deck_type || '', 80).trim() || null;
  const lockAgency = shouldLockAgency10(extractJson, options);
  const requestedDeckType = lockAgency ? 'ad_agency' : requestedDeckTypeRaw;
  const requestedSlides = lockAgency ? 10 : nSlides;

  const schema = buildDeckSchema();
  const system = buildAssembleSystemPrompt({ vibe, audience, language, requestedDeckType, requestedSlides, voiceProfile });
  const user = buildAssembleUserPrompt(extractJson, narrativeJson, messagingMap);

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.7
  });

  const content = completion.choices?.[0]?.message?.content;
  if (!content) throw new Error('No plan returned from model.');
  return JSON.parse(content);
}

async function assembleDeckWithGemini(extractJson, narrativeJson, messagingMap, options = {}) {
  const { gemini_text: defaultGeminiTextModel } = getModels();

  const nSlides = clampInt(options.nSlides ?? 10, 5, SAFE_MAX_SLIDES, 10);
  const voiceProfile = resolveVoiceProfile(options);
  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);
  const requestedDeckTypeRaw = asStr(options.deckType || options.deck_type || '', 80).trim() || null;
  const lockAgency = shouldLockAgency10(extractJson, options);
  const requestedDeckType = lockAgency ? 'ad_agency' : requestedDeckTypeRaw;
  const requestedSlides = lockAgency ? 10 : nSlides;

  const schema = buildDeckSchema();
  const system = buildAssembleSystemPrompt({ vibe, audience, language, requestedDeckType, requestedSlides, voiceProfile });
  const user = buildAssembleUserPrompt(extractJson, narrativeJson, messagingMap);

  return geminiGenerateJson({
    model: defaultGeminiTextModel,
    system,
    user,
    jsonSchema: schema.schema,
    temperature: 0.7
  });
}

async function editDeckWithOpenAI(extractJson, narrativeJson, messagingMap, deckPlan, options = {}) {
  const client = getOpenAIClient();
  const { text: model } = getModels();

  const voiceProfile = resolveVoiceProfile(options);
  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);

  const schema = buildDeckSchema();
  const system = buildEditSystemPrompt({
    vibe,
    audience,
    language,
    voiceProfile,
    requestedDeckType: shouldLockAgency10(extractJson, options) ? 'ad_agency' : null
  });
  const user = buildEditUserPrompt(extractJson, narrativeJson, messagingMap, deckPlan);

  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.35
  });

  const content = completion.choices?.[0]?.message?.content;
  if (!content) throw new Error('No edited deck_plan returned from model.');
  return JSON.parse(content);
}

async function editDeckWithGemini(extractJson, narrativeJson, messagingMap, deckPlan, options = {}) {
  const { gemini_text: defaultGeminiTextModel } = getModels();

  const voiceProfile = resolveVoiceProfile(options);
  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);

  const schema = buildDeckSchema();
  const system = buildEditSystemPrompt({
    vibe,
    audience,
    language,
    voiceProfile,
    requestedDeckType: shouldLockAgency10(extractJson, options) ? 'ad_agency' : null
  });
  const user = buildEditUserPrompt(extractJson, narrativeJson, messagingMap, deckPlan);

  return geminiGenerateJson({
    model: defaultGeminiTextModel,
    system,
    user,
    jsonSchema: schema.schema,
    temperature: 0.35
  });
}

// -------- Public API --------

/**
 * Two-pass planning (extract → assemble). Default ON.
 * options.provider: 'openai' | 'gemini'
 * options.twoPass: boolean (default true)
 */
export async function planDeck(briefText, options = {}) {
  const provider = (options.provider || 'openai').toString().toLowerCase();
  const outlineSignals = extractExplicitOutlineSignals(briefText);

  const twoPass = options.twoPass !== false;

  if (!twoPass) {
    // Fallback to the old one-pass behavior (still works, but less reliable).
    // We keep it for debugging.
    return planDeckOnePass(briefText, options);
  }

  const extract = provider === 'gemini'
    ? await extractBriefWithGemini(briefText, options)
    : await extractBriefWithOpenAI(briefText, options);

  // Attach explicit slide-by-slide outline hints if present (used to expand naturally; not hard-enforced)
  if (outlineSignals?.outline?.length) {
    extract.explicit_outline = outlineSignals.outline;
  }
  if (outlineSignals?.range) {
    extract.slide_count_range_hint = outlineSignals.range;
  }

  // New: narrative blueprint step (enforces journey + cohesion)
  const narrative = provider === 'gemini'
    ? await planNarrativeWithGemini(extract, options)
    : await planNarrativeWithOpenAI(extract, options);

  // Hard-lock agency creative decks to the requested 10-slide structure.
  const narrativeLocked = shouldLockAgency10(extract, options)
    ? lockNarrativeToAgency10(narrative, extract)
    : narrative;

  // New: copy constraints step ("agency brain": anchors, section concepts, bridges)
  const messaging = provider === 'gemini'
    ? await planMessagingWithGemini(extract, narrativeLocked, options)
    : await planMessagingWithOpenAI(extract, narrativeLocked, options);

  const draftPlan = provider === 'gemini'
    ? await assembleDeckWithGemini(extract, narrativeLocked, messaging, options)
    : await assembleDeckWithOpenAI(extract, narrativeLocked, messaging, options);

  const useEditorPass = options.editorPass !== false;
  const plan = useEditorPass
    ? (provider === 'gemini'
      ? await editDeckWithGemini(extract, narrativeLocked, messaging, draftPlan, options)
      : await editDeckWithOpenAI(extract, narrativeLocked, messaging, draftPlan, options))
    : draftPlan;

  // Final enforcement: exactly 10 slides in the required order for agency creative decks.
let planLocked = shouldLockAgency10(extract, options)
  ? lockDeckPlanToAgency10(plan, extract, options)
  : plan;

// Agency-only: refine concept line quality (no structure changes)
if (shouldLockAgency10(extract, options)) {
  planLocked = provider === 'gemini'
    ? await refineAgencyConceptWithGemini(briefText, extract, narrativeLocked, messaging, planLocked, options)
    : await refineAgencyConceptWithOpenAI(briefText, extract, narrativeLocked, messaging, planLocked, options);
}

// Attach extract for debugging / UI JSON edit (safe)
  planLocked._extract = extract;
  planLocked._narrative = narrativeLocked;
  planLocked._messaging = messaging;
  return planLocked;
}

// ----- One-pass (legacy) -----

function buildLegacyDeckSchema() {
  return {
    name: 'deck_plan_legacy',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        deck_type: { type: 'string', enum: DECK_TYPES },
        deck_title: { type: 'string' },
        deck_subtitle: { type: 'string' },
        recommended_slide_count: { type: 'integer', minimum: 5, maximum: SAFE_MAX_SLIDES },
        theme: {
          type: 'object',
          additionalProperties: false,
          properties: {
            vibe: { type: 'string' },
            primary_color: { type: 'string' },
            secondary_color: { type: 'string' },
            font_heading: { type: 'string' },
            font_body: { type: 'string' }
          },
          required: ['vibe', 'primary_color', 'secondary_color', 'font_heading', 'font_body']
        },
        slides: {
          type: 'array',
          minItems: 5,
          maxItems: SAFE_MAX_SLIDES,
          items: {
            type: 'object',
            additionalProperties: false,
            properties: {
              kind: { type: 'string' },
              layout: { type: 'string', enum: ['hero', 'split', 'full_bleed', 'quote', 'stats', 'two_column'] },
              title: { type: 'string' },
              subtitle: { type: 'string' },
              bullets: { type: 'array', items: { type: 'string' }, maxItems: 6 },
              stat: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: { value: { type: 'string' }, label: { type: 'string' } },
                required: ['value', 'label']
              },
              quote: {
                type: ['object', 'null'],
                additionalProperties: false,
                properties: { text: { type: 'string' }, attribution: { type: 'string' } },
                required: ['text', 'attribution']
              },
              image_prompt: { type: 'string' },
              speaker_notes: { type: 'string' }
            },
            required: ['kind', 'layout', 'title', 'subtitle', 'bullets', 'stat', 'quote', 'image_prompt', 'speaker_notes']
          }
        }
      },
      required: ['deck_type', 'deck_title', 'deck_subtitle', 'recommended_slide_count', 'theme', 'slides']
    }
  };
}

function buildLegacySystemPrompt({ vibe, audience, language, requestedDeckType, requestedSlides }) {
  return [
    `You are a senior creative director and presentation strategist.`,
    `Create a slide-by-slide plan. Output MUST match schema strictly.`,
    `Use concise copy, strong narrative, layout variety.`,
    `Infer deck_type (or use requested).`,
    requestedDeckType ? `Requested deck type: ${requestedDeckType}.` : `No explicit deck type; infer from brief.`,
    requestedSlides ? `Target slide count: ${requestedSlides}.` : `Choose 5–18 slides.`,
    `Language: ${language}. Audience: ${audience}. Vibe: ${vibe}.`
  ].join(' ');
}

async function planDeckOnePass(briefText, options = {}) {
  const provider = (options.provider || 'openai').toString().toLowerCase();
  const nSlides = clampInt(options.nSlides ?? 10, 5, SAFE_MAX_SLIDES, 10);
  const vibe = asStr(options.vibe || 'Modern, premium', 120);
  const audience = asStr(options.audience || 'general', 120);
  const language = asStr(options.language || 'English', 80);
  const requestedDeckType = asStr(options.deckType || options.deck_type || '', 80).trim() || null;

  const schema = buildLegacyDeckSchema();
  const system = buildLegacySystemPrompt({ vibe, audience, language, requestedDeckType, requestedSlides: nSlides });
  const user = buildUserPrompt(briefText);

  if (provider === 'gemini') {
    return geminiGenerateJson({
      model: getModels().gemini_text,
      system,
      user,
      jsonSchema: schema.schema,
      temperature: 0.7
    });
  }

  const client = getOpenAIClient();
  const { text: model } = getModels();
  const completion = await client.chat.completions.create({
    model,
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    response_format: { type: 'json_schema', json_schema: schema },
    temperature: 0.7
  });
  const content = completion.choices?.[0]?.message?.content;
  if (!content) throw new Error('No plan returned from model.');
  return JSON.parse(content);
}

// -------- Normalization (safe for renderer) --------

function defaultTheme() {
  return {
    vibe: 'Modern, premium',
    primary_color: '#0B0F1A',
    secondary_color: '#2A7FFF',
    font_heading: 'Aptos Display',
    font_body: 'Aptos'
  };
}

function defaultSlide(idx) {
  return {
    kind: idx === 0 ? 'cover' : 'content',
    layout: idx === 0 ? 'hero' : 'split',
    section: '',
    setup_line: '',
    takeaway: '',
    bridge_line: '',
    title: idx === 0 ? 'Title' : `Slide ${idx + 1}`,
    subtitle: '',
    bullets: [],
    stat: null,
    quote: null,
    agenda_items: null,
    cards: null,
    timeline_items: null,
    kpis: null,
    status_items: null,
    table: null,
    pricing: null,
    matrix: null,
    steps: null,
    people: null,
    logo_items: null,
    cta: null,
    swot: null,
    funnel: null,
    now_next_later: null,
    okrs: null,
    case_study: null,
    diagram: null,
    icons: null,
    chart: null,
    org_chart: null,
    faq: null,
    image_prompt: 'Abstract premium background related to the slide topic',
    speaker_notes: ''
  };
}

function inferLayoutFromKind(kind) {
  const k = (kind || '').toString().toLowerCase();
  if (k.includes('cover')) return 'hero';
  if (k.includes('agenda')) return 'agenda';
  if (k.includes('section')) return 'section_header';
  if (k.includes('timeline') || k.includes('roadmap')) return 'timeline';
  if (k.includes('swot')) return 'swot';
  if (k.includes('funnel')) return 'funnel';
  if (k.includes('now') && k.includes('next') && k.includes('later')) return 'now_next_later';
  if (k.includes('okr') || k.includes('okrs')) return 'okr';
  if (k.includes('case study') || k.includes('casestudy')) return 'case_study';
  if (k.includes('chart') || k.includes('trend')) return 'chart_line';
  if (k.includes('diagram') || k.includes('flow')) return 'diagram';
  if (k.includes('org') || k.includes('organisation') || k.includes('organization')) return 'org_chart';
  if (k.includes('faq') || k.includes('q&a')) return 'faq';
  if (k.includes('kpi') || k.includes('metrics') || k.includes('dashboard')) return 'kpi_dashboard';
  if (k.includes('status') || k.includes('ryg') || k.includes('traffic')) return 'traffic_light';
  if (k.includes('pricing') || k.includes('package')) return 'pricing';
  if (k.includes('comparison') || k.includes('matrix')) return 'comparison_matrix';
  if (k.includes('process') || k.includes('steps')) return 'process_steps';
  if (k.includes('team')) return 'team_grid';
  if (k.includes('clients') || k.includes('logos')) return 'logo_wall';
  if (k.includes('pillar') || k.includes('infographic')) return 'infographic_3';
  if (k.includes('cta') || k.includes('next steps') || k.includes('contact')) return 'cta';
  return null;
}

export function normalizePlan(plan, options = {}) {
  const requestedSlides = clampInt(
    options.nSlides ?? plan?.recommended_slide_count ?? (Array.isArray(plan?.slides) ? plan.slides.length : 10),
    5,
    SAFE_MAX_SLIDES,
    10
  );

  const deckType = asStr(plan.deck_type || options.deckType || options.deck_type || plan?._extract?.deck_type_suggestion || 'other', 80)
    .toLowerCase()
    .trim();

  const theme = { ...defaultTheme(), ...(plan.theme || {}) };
  if (options.deckStyle && !theme.deck_style) theme.deck_style = options.deckStyle;

  let slides = (plan.slides || []).map((s, idx) => {
    const base = defaultSlide(idx);
    const kind = asStr(s?.kind || base.kind, 60);
    const inferred = inferLayoutFromKind(kind);
    const layoutRaw = asStr(s?.layout || inferred || base.layout, 40);
    const layout = SLIDE_LAYOUTS.includes(layoutRaw) ? layoutRaw : base.layout;

    const safe = {
      ...base,
      kind,
      layout,
      title: asStr(s?.title || base.title, 140),
      subtitle: asStr(s?.subtitle || '', 240),
      bullets: Array.isArray(s?.bullets) ? s.bullets.slice(0, 8).map(v => asStr(v, 180)).filter(Boolean) : [],
      stat: s?.stat ?? null,
      quote: s?.quote ?? null,
      agenda_items: Array.isArray(s?.agenda_items) ? s.agenda_items.slice(0, 10).map(v => asStr(v, 120)).filter(Boolean) : null,
      cards: Array.isArray(s?.cards) ? s.cards.slice(0, 6).map(c => ({
        title: asStr(c?.title, 80),
        body: asStr(c?.body, 220),
        tag: asStr(c?.tag, 40)
      })) : null,
      timeline_items: Array.isArray(s?.timeline_items) ? s.timeline_items.slice(0, 12).map(t => ({
        date_or_phase: asStr(t?.date_or_phase, 40),
        label: asStr(t?.label, 80),
        detail: asStr(t?.detail, 140)
      })) : null,
      kpis: Array.isArray(s?.kpis) ? s.kpis.slice(0, 8).map(k => ({
        label: asStr(k?.label, 60),
        value: asStr(k?.value, 40),
        delta: asStr(k?.delta, 30)
      })) : null,
      status_items: Array.isArray(s?.status_items) ? s.status_items.slice(0, 12).map(it => ({
        item: asStr(it?.item, 90),
        status: ['red','yellow','green'].includes((it?.status || '').toString().toLowerCase()) ? (it.status || '').toString().toLowerCase() : 'yellow',
        owner: asStr(it?.owner, 40),
        eta: asStr(it?.eta, 30),
        blocker: asStr(it?.blocker, 120)
      })) : null,
      table: (s?.table && Array.isArray(s.table.headers) && Array.isArray(s.table.rows)) ? {
        headers: s.table.headers.slice(0, 6).map(h => asStr(h, 40)),
        rows: s.table.rows.slice(0, 12).map(r => (Array.isArray(r) ? r.slice(0, 6).map(v => asStr(v, 40)) : [])).filter(r => r.length >= 2)
      } : null,
      pricing: s?.pricing ?? null,
      matrix: s?.matrix ?? null,
      steps: Array.isArray(s?.steps) ? s.steps.slice(0, 8).map(st => ({ title: asStr(st?.title, 70), detail: asStr(st?.detail, 160) })) : null,
      people: Array.isArray(s?.people) ? s.people.slice(0, 10).map(p => ({ name: asStr(p?.name, 60), role: asStr(p?.role, 60), bio: asStr(p?.bio, 200) })) : null,
      logo_items: Array.isArray(s?.logo_items) ? s.logo_items.slice(0, 30).map(v => asStr(v, 40)).filter(Boolean) : null,
      cta: s?.cta ?? null,
      swot: s?.swot ?? null,
      funnel: Array.isArray(s?.funnel) ? s.funnel.slice(0, 7).map(st => ({ label: asStr(st?.label, 60), value: asStr(st?.value, 40) })) : null,
      now_next_later: s?.now_next_later ?? null,
      okrs: Array.isArray(s?.okrs) ? s.okrs.slice(0, 5).map(o => ({ objective: asStr(o?.objective, 120), key_results: Array.isArray(o?.key_results) ? o.key_results.slice(0, 6).map(v => asStr(v, 160)).filter(Boolean) : [] })) : null,
      case_study: s?.case_study ?? null,
      diagram: s?.diagram ? {
        code: asStr(s.diagram.code || '', 4000),
        theme: asStr(s.diagram.theme || '', 40)
      } : null,
      icons: Array.isArray(s?.icons) ? s.icons.slice(0, 6).map(ic => ({
        name: asStr(ic?.name || '', 120),
        label: asStr(ic?.label || '', 80)
      })) : null,
      chart: s?.chart ? {
        chart_type: asStr(s.chart.chart_type || '', 20),
        labels: Array.isArray(s.chart.labels) ? s.chart.labels.slice(0, 10).map(v => asStr(v, 80)) : [],
        values: Array.isArray(s.chart.values) ? s.chart.values.slice(0, 10).map(v => Number(v)) : [],
        value_suffix: asStr(s.chart.value_suffix || '', 20),
        spec: s.chart.spec ?? null
      } : null,
      org_chart: s?.org_chart ?? null,
      faq: Array.isArray(s?.faq) ? s.faq.slice(0, 8).map(it => ({ q: asStr(it?.q, 140), a: asStr(it?.a, 220) })) : null,
      image_prompt: asStr(s?.image_prompt ?? base.image_prompt, 800),
      speaker_notes: asStr(s?.speaker_notes ?? '', 1600)
    };

    // For data-heavy slides, default to no image unless explicitly provided.
    if (['kpi_dashboard','traffic_light','table','pricing','comparison_matrix','process_steps','team_grid','logo_wall','swot','funnel','now_next_later','okr','chart_bar','chart_line','diagram','org_chart','faq','case_study','appendix','infographic_3'].includes(layout)) {
      const p = (safe.image_prompt || '').trim();
      if (!p || /^none$/i.test(p)) safe.image_prompt = 'NONE';
    }

    return safe;
  });

  if (deckType !== 'ad_agency') {
    const hasQuote = slides.some(s => s.layout === 'quote' || (s.kind || '').toLowerCase().includes('quote'));
    const hasInfographic = slides.some(s => s.layout === 'infographic_3' || (s.kind || '').toLowerCase().includes('pillar'));

    if (!hasInfographic) {
      slides.splice(Math.min(2, slides.length), 0, {
        kind: 'infographic_summary',
        layout: 'infographic_3',
        title: 'Speak Locally. Resonate Nationally.',
        subtitle: 'Our strategy turns complexity into three simple pillars:',
        cards: [
          { title: 'Hyper-Local Pride', body: 'Celebrate unique identity and local cues.', tag: 'pin' },
          { title: 'Universal Invitation', body: 'Unify the campaign under one idea.', tag: 'signal' },
          { title: 'Simplicity & Clarity', body: 'Make the call-to-action unforgettable.', tag: 'hashtag' }
        ],
        image_prompt: 'NONE',
        speaker_notes: ''
      });
    }

    if (!hasQuote) {
      const quoteText = asStr(plan?._extract?.objective || plan?._extract?.source_summary || 'A single, memorable idea beats a dozen scattered messages.', 260);
      slides.splice(Math.max(slides.length - 1, 1), 0, {
        kind: 'quote',
        layout: 'quote',
        title: 'Key Thought',
        subtitle: '',
        quote: { text: quoteText, attribution: '' },
        image_prompt: 'Abstract premium background, subtle texture, no text',
        speaker_notes: ''
      });
    }
  }

  while (slides.length > requestedSlides) {
    let removed = false;
    for (let i = slides.length - 1; i >= 0; i--) {
      const layout = (slides[i]?.layout || '').toString().toLowerCase();
      if (layout === 'quote' || layout === 'infographic_3') continue;
      slides.splice(i, 1);
      removed = true;
      break;
    }
    if (!removed) break;
  }

  slides = slides.slice(0, requestedSlides);

  return {
    deck_type: deckType,
    deck_title: asStr(plan.deck_title || plan?._extract?.title || 'Untitled Deck', 140),
    deck_subtitle: asStr(plan.deck_subtitle || plan?._extract?.subtitle || '', 200),
    recommended_slide_count: clampInt(plan.recommended_slide_count ?? requestedSlides, 5, SAFE_MAX_SLIDES, requestedSlides),
    theme,
    slides,
    brand_logo: plan?.brand_logo || plan?.theme?.brand_logo || null,
    _extract: plan?._extract || null
  };
}
