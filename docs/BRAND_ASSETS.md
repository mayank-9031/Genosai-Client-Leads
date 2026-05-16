# GenosAI — Brand Assets & Design System

> Single source of truth for designers and AI design tools (Claude design, Canva, Figma, slide/deck generators, social carousel generators). Everything below is derived from the production website at https://genosai.tech.

**Brand:** GenosAI
**Domain:** genosai.tech
**Contact:** hello@genosai.tech
**Socials:** LinkedIn `/company/genosai` · X `@AiGenos81577`
**HQ:** India · **Markets:** USA, UK, UAE, India, Australia

---

## 1. Brand Identity

### 1.1 What we are
GenosAI is a **global AI automation agency** that builds **production-grade AI systems** — chatbots, voice AI agents, workflow automation, lead qualification engines, and bespoke enterprise AI — for businesses that want to eliminate manual operations and scale.

### 1.2 Positioning statement
> We build AI systems that automate your sales, outreach, support, social media, and business operations.

### 1.3 Brand essence (one-liners)
- **Tagline (rotating):** *We Build AI Systems That Automate Your [Sales / Outreach / Support / Social Media / Business Operations].*
- **Manifesto line:** *We're engineers, not salespeople.*
- **Closer line:** *We don't chase hype. We ship systems that work.*
- **Promise:** *Production-grade from day one.*

### 1.4 Voice & tone
| Dimension | Direction |
|---|---|
| **Persona** | Senior engineer, calm, technically confident, outcome-led |
| **Tone** | Direct, plain English, zero hype, lightly understated |
| **Avoid** | Buzzword stacking ("synergize", "revolutionize"), exclamation marks, emoji-led copy, vague AI jargon ("agentic future") |
| **Prefer** | "We ship", "We build", "Engineered", "Eliminate manual ops", "Integrates with your stack", "Runs in production from day one" |
| **Sentence rhythm** | Short. Concrete. Then one longer line that lands the proof. |

### 1.5 Proof points (use as social proof in any deck/post)
- **23+ clients** across 5 countries
- **50+ projects** shipped
- **98% client retention**
- **4–12 week** average delivery
- Markets: 🇺🇸 USA · 🇬🇧 UK · 🇦🇪 UAE · 🇮🇳 India · 🇦🇺 Australia

---

## 2. Logo

The brand mark is a **wordmark**, not an icon. There is no separate symbol/avatar mark.

### 2.1 Wordmark
- **Text:** `GenosAI` (one word, capital G, capital A — never "Genos AI", "GENOSAI", or "genosai")
- **Footer treatment:** `Genos` + `AI` set in the same color (white on dark). Avoid color-splitting the two halves.
- **Typeface:** Instrument Serif, Regular (400)
- **Letter-spacing:** −0.02em (slightly tight)
- **Default color on dark:** `#FFFFFF`
- **Default color on light:** `#0A0A0F`

### 2.2 Sizing reference
| Use | Size |
|---|---|
| Nav / footer | 1.5 rem (24 px) |
| Slide title-card / poster | 96–160 px cap-height |
| Instagram carousel cover | 180–240 px cap-height on a 1080² canvas |
| Favicon / app icon | Use a square crop of the capital `G` in Instrument Serif, white on `#0A0A0F` |

### 2.3 Clear space & don'ts
- **Clear space:** at least the cap-height of the `G` on every side.
- **Do not:** add drop shadows, outline, gradient fills, italicize, condense, stretch, or place over busy imagery without a `#0A0A0F` ≥ 60% scrim.
- **Do not:** recreate the wordmark in a different typeface — Instrument Serif is the lockup.

---

## 3. Color System

The site lives in a **dark, near-black canvas with warm off-white type**, lifted by a **single violet accent** used sparingly on eyebrows and ambient glows. White is the primary brand color — violet is a spice, not a base.

### 3.1 Core palette

| Token | Hex | Role |
|---|---|---|
| **bg-dark** | `#0A0A0F` | Primary canvas (≈99% of surface area) |
| **bg-accent** | `#111118` | Elevated surface / card on dark |
| **bg-light** | `#F0ECE4` | Light/cream alt canvas (rare; print, light-mode decks) |
| **text-on-dark** | `#F0EDE8` | Body text on dark (warm off-white, not pure white) |
| **text-on-light** | `#1A1A1A` | Body text on light canvas |
| **white** | `#FFFFFF` | Headlines, primary buttons, logo |
| **soft-grey** | `#AAAAAA` | Muted text / dividers |

### 3.2 Accent

| Token | Hex | Role |
|---|---|---|
| **violet-primary** | `#6D28D9` | Ambient hero glow / dither shader tint |
| **violet-eyebrow** | `#A78BFA` (used at ~70% opacity) | Section eyebrow labels only |

Use violet **only** for: section eyebrows ("WHO WE ARE", "AI AUTOMATION AGENCY"), soft background glows behind hero/CTA, occasional 1-pixel highlight. **Never** for body text, body buttons, or large color blocks.

### 3.3 Functional opacities (white on `#0A0A0F`)
The site builds depth almost entirely from white at varying opacities. Reuse these exactly:

| Opacity | Use |
|---|---|
| 100% | Primary headlines, primary button bg, logo |
| 70% | Active body / interactive label |
| 55% | Hero subhead body |
| 50% | Body paragraphs |
| 40% | Footer body, secondary labels |
| 30% | Footer headings, dim labels |
| 25% | Copyright, fine print |
| 20% | Secondary button border |
| 10% | Card icon background |
| 6% | Card border (`white/[0.06]`) |
| 3% | Card surface tint (`white/[0.03]`) |
| 0% (transparent) | Dot pattern fade-out |

### 3.4 Status / chart colors
The site does not use saturated semantic colors. For decks/dashboards, use neutral greyscale steps for charts and reserve violet (`#6D28D9`) for "active" or "highlight" series only.

---

## 4. Typography

Two fonts. One serif for display, one sans for body. Don't introduce a third.

### 4.1 Type families
| Role | Family | Weights | Source |
|---|---|---|---|
| **Display / Headlines** | **Instrument Serif** | 400 (Regular), 400 Italic | Google Fonts |
| **Body / UI** | **Satoshi** | 400, 500, 600, 700, 800 | Fontshare (`api.fontshare.com`) |
| **Mono** (rare; code/tables) | Geist Mono | 400, 500 | Vercel |

If Satoshi is unavailable in a deck tool, fall back to: **Inter → SF Pro → system-ui sans-serif**.
If Instrument Serif is unavailable, fall back to: **Cormorant Garamond → EB Garamond → Georgia**.

### 4.2 Type scale (web → translates 1:1 to slides/social)

| Style | Family | Size | Weight | Line-height | Tracking | Color |
|---|---|---|---|---|---|---|
| **Display XL** (hero) | Instrument Serif | clamp(2.5rem, 6vw, 5.5rem) → ~88 px | 400 | 0.95 | −0.03em | white→neutral-400 vertical gradient |
| **Display L** (section H2) | Instrument Serif | clamp(2rem, 4vw, 3.5rem) → ~56 px | 400 | 1.05 | −0.02em | `#F0EDE8` |
| **Display M** (stat numbers) | Instrument Serif | 2.5rem (40 px) | 400 | 1.0 | −0.01em | white |
| **Display Italic** (pull quote) | Instrument Serif Italic | 1rem (16 px) | 400 | 1.7 | 0 | `#F0EDE8` |
| **Body L** (hero subhead) | Satoshi | 1.05–1.15rem (17–18 px) | 400 | 1.7 | 0 | white/55 |
| **Body M** (paragraphs) | Satoshi | 0.95rem (15 px) | 400 | 1.7 | 0 | white/50 |
| **Body S** (footer) | Satoshi | 0.85–0.88rem (14 px) | 400 | 1.6 | 0 | white/40–50 |
| **Eyebrow / kicker** | Satoshi | 0.7rem (11 px) | 600 | 1.0 | 0.2em | violet-400/70, **UPPERCASE** |
| **Section heading mini** | Satoshi | 0.7rem (11 px) | 600 | 1.0 | 0.15em | white/30, **UPPERCASE** |
| **Button label** | Satoshi | 0.9rem (14 px) | 600 | 1.0 | 0.06em | varies, **UPPERCASE** |
| **Caption / copyright** | Satoshi | 0.75rem (12 px) | 400 | 1.4 | 0 | white/25 |

### 4.3 Headline rules
- Headlines **always** use Instrument Serif at weight 400 — never bold a serif headline.
- Use a **vertical white→light-grey gradient** (`#FAFAFA → #A3A3A3`) on hero-tier headlines for a subtle metallic feel. Body headlines stay solid `#F0EDE8`.
- Use **negative letter-spacing** (−0.02em to −0.03em) on serif display sizes.
- Eyebrows ABOVE every section, in the violet 11px UPPERCASE 0.2em treatment.

---

## 5. Components & UI Patterns

### 5.1 Buttons

**Primary (CTA)**
- Background: `#FFFFFF`
- Text: `#0A0A0F`, Satoshi 600, 0.9rem, UPPERCASE, tracking 0.06em
- Padding: 16 px vertical · 32 px horizontal (`py-4 px-8`)
- Radius: **fully pill** (`rounded-full`)
- Hover: lift 2 px + soft white glow `0 8px 40px rgba(255,255,255,0.15)`
- Example labels: `Book a Free Strategy Call` · `Get Started`

**Secondary (ghost)**
- Background: transparent
- Border: 1 px `rgba(255,255,255,0.20)`
- Text: `rgba(255,255,255,0.70)`, same type spec as primary
- Hover: border opacity → 40%, text → 100%
- Example: `See Our Services` · `Learn More`

### 5.2 Cards / surfaces
- Radius: **`rounded-2xl`** (16 px) — never sharp corners
- Border: 1 px `rgba(255,255,255,0.06)`
- Surface: `linear-gradient(135deg, rgba(255,255,255,0.03) 0%, transparent 100%)`
- Optional: subtle dot grid at 3% opacity, 24 px spacing, on top of the card
- Optional: animated glow / proximity-aware border (decorative only)

### 5.3 Section structure
Every section follows this rhythm — reuse it for slide layouts and Instagram carousels:

1. **Eyebrow** (violet 11 px UPPERCASE 0.2em tracking)
2. **Headline** (Instrument Serif display)
3. **One-paragraph subhead** (Satoshi, white/50, line-height 1.7)
4. **Optional visual** (right column on desktop, below on mobile)

### 5.4 Backgrounds & visual motifs
| Motif | When to use |
|---|---|
| **Solid `#0A0A0F`** | Default canvas |
| **Soft radial vignette** | Hero / cover slide — `radial-gradient(ellipse at center, transparent 0%, transparent 40%, rgba(10,10,15,0.55) 100%)` |
| **Violet dither / shader** | Hero only, ~18% opacity over solid bg, subtle warp |
| **Dot grid** | Card surfaces, 24 px spacing, 3% opacity white dots |
| **World map with arc lines** | "Global reach" / regions slide; lines drawn in pure white `#FFFFFF` at thin stroke |
| **Spotlight cursor / glow orb** | Hero accent, soft white at low opacity |
| **Etheral shadow** | Site-wide ambient backdrop |

**Avoid:** glassmorphism with strong blur, neon, gradients with more than 2 stops, stock-photo collages.

### 5.5 Iconography
- **Library:** [Lucide](https://lucide.dev) (already used: `Globe2`, `Users`, `Zap`, `Linkedin`, `Twitter`).
- **Style:** stroke 1.5–2 px, white at 40% opacity by default, in a 36 px rounded-square chip with `rgba(255,255,255,0.10)` background.
- **Avoid:** filled / multi-color / 3D icon sets.

### 5.6 Imagery
- Service tiles use **abstract AI imagery** — neural meshes, blue/purple orbs, glass-like organic forms. Never people-stock.
- All photographic imagery sits on dark backgrounds; tint with a subtle `#0A0A0F` overlay if needed for legibility.
- Map graphics use thin pure-white arcs on a dotted globe.

---

## 6. Layout Specs

### 6.1 Web (reference)
- Section padding: vertical 64–96 px (`py-16 md:py-24`), horizontal 5vw
- Max content width: 1280 px (`max-w-7xl`) for full sections, 1024 px (`max-w-5xl`) for hero copy, 720 px for body paragraphs
- Grid: 12-column on desktop, single column on mobile

### 6.2 Slide deck (16:9, 1920×1080)
| Element | Spec |
|---|---|
| Outer padding | 96 px |
| Title | Instrument Serif 96–120 px, line-height 1.0, tracking −0.02em, white→grey gradient |
| Eyebrow | Satoshi 600, 18 px, UPPERCASE, tracking 0.2em, violet `#A78BFA` @ 70% |
| Body | Satoshi 400, 28 px, line-height 1.5, `rgba(255,255,255,0.6)` |
| Footer / page number | Satoshi 400, 14 px, white @ 25% |

### 6.3 Instagram carousel (1080×1080 or 1080×1350)
| Element | Spec |
|---|---|
| Outer padding | 80 px |
| Title | Instrument Serif 84–110 px, line-height 0.95 |
| Eyebrow | Satoshi 600, 22 px, tracking 0.2em, UPPERCASE, violet @ 70% |
| Body | Satoshi 400, 36 px, line-height 1.5, white @ 60% |
| Wordmark (bottom-right) | Instrument Serif `GenosAI`, 32 px, white @ 80% |
| Page indicator (e.g., 02 / 06) | Satoshi 600, 20 px, white @ 30%, tracking 0.15em |

### 6.4 Story / Reel cover (1080×1920)
- Same type system, body sizes ~10–15% larger
- Wordmark anchored top-left, CTA pill anchored bottom-center

---

## 7. Service Lineup (use as content blocks)

Use these short forms verbatim in decks/social:

| # | Service | One-line |
|---|---|---|
| 01 | **AI Chatbot Development** | Conversational AI assistants tuned to your data and brand voice. |
| 02 | **Voice AI Agents** | Outbound calling, lead qualification, meeting booking — at scale. |
| 03 | **Workflow Automation** | End-to-end ops: CRMs, ERPs, spreadsheets, dashboards. No more manual data entry. |
| 04 | **AI Marketing & Outreach** | Automated sequences, lead scoring, behavior-triggered campaigns. |
| 05 | **AI Lead Qualification** | Score, qualify and route inbound leads in real time. |
| 06 | **Custom AI Development** | Bespoke enterprise systems — vision, NLP, recommendation, decisioning. |

### 7.1 Tech stack (logos available on request)
- **AI models:** OpenAI · Anthropic Claude · Google Gemini · Meta LLaMA · Mistral · DeepSeek · ElevenLabs · Whisper
- **Automation:** Make · n8n · Zapier · LangChain · LangGraph · CrewAI · Twilio · Retell AI
- **Infra:** Next.js · React · Node.js · Python · FastAPI · PostgreSQL · Supabase · Vercel · AWS · Docker

---

## 8. Copy Bank (drop-in lines)

Use these as headline / hook seeds for decks and social:

- *We build AI systems that automate your business operations.*
- *Engineers, not salespeople.*
- *Production-grade from day one.*
- *We don't chase hype. We ship systems that work.*
- *50+ projects. 5 countries. 98% retention.*
- *Custom AI that integrates with the stack you already use.*
- *From repetitive tasks to autonomous workflows — in 4 to 12 weeks.*
- *AI that runs your business, so you can grow it.*

---

## 9. Quick design checklist (for any new asset)

Before you ship a deck, post, or carousel, verify:

- [ ] Canvas is `#0A0A0F`, not pure black `#000000`
- [ ] All headlines are **Instrument Serif 400** (never bold serif)
- [ ] All body/UI is **Satoshi**
- [ ] Section opens with a **violet UPPERCASE eyebrow** at 0.2em tracking
- [ ] No more than **one violet accent** per slide
- [ ] Body color is **white at 50–60% opacity**, not pure white
- [ ] Buttons are **pill-shaped** (`rounded-full`), white-on-dark for primary
- [ ] Cards use `rounded-2xl` and a 6%-white border — never sharp corners
- [ ] No emojis, no exclamation marks, no buzzwords
- [ ] Logo wordmark `GenosAI` appears once per asset, never restyled
- [ ] If imagery is used, it's abstract AI / neural / orb — never people-stock

---

## 10. File reference (where this comes from in code)

| Spec | File |
|---|---|
| Color tokens, animations | [src/app/globals.css](src/app/globals.css) |
| Fonts, metadata, theme color | [src/app/layout.tsx](src/app/layout.tsx) |
| Hero copy / type scale | [src/components/HeroSection.tsx](src/components/HeroSection.tsx) |
| Section pattern, stats, voice | [src/components/sections/About.tsx](src/components/sections/About.tsx) |
| Service line-up | [src/components/sections/ServicesGrid.tsx](src/components/sections/ServicesGrid.tsx) |
| Tech stack | [src/components/sections/TechStack.tsx](src/components/sections/TechStack.tsx) |
| Logo treatment, footer, social | [src/components/sections/Footer.tsx](src/components/sections/Footer.tsx) |
