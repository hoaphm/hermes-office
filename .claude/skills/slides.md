---
name: slides
description: >-
  Build premium PRESENTED slide decks — classic one-slide-at-a-time slides you advance with a clicker — as a Vite + React app. Keeps Slidev navigation UI (floating glass dock + thumbnail rail + click-builds + presenter) but slides are RESPONSIVE React (reflow to any screen, no fixed canvas) and fully interactive. Use for decks you'll PRESENT (talk over, projector, screen-share) when you want web interactivity and responsiveness without Slidev's constraints. The repo IS the running app (Vite + React) — you install it in place, theme tokens, and AUTHOR an original deck from the user's real topic/brand.
---

# Slides — premium, responsive, React presentation engine

Bolt Classic **paged** slides (advance one at a time, present over them) — but rebuilt in **Vite + React** so each slide is a **responsive web layout** instead of a fixed 1080×607 canvas clip. It keeps the Slidev UI you liked (the floating dock + thumbnail rail + click-builds + presenter) and adds real web interactivity.

Two halves — keep them separate:

- **The engine + UI are pre-built** — the repo is a complete, runnable app: paged engine (`src/deck/`), dock/rail chrome, section components (`src/components/`), shared CSS. **Leave the engine as-is. Never regenerate it.** This part you liked; it must look/behave identically.
- **The content is authored fresh, every time** — slides (topic, structure, copy, visuals, theme) are designed from scratch for *this* request.

**Theme surface.** All color, type, radius, depth, motion live in the `:root` token vocabulary in `src/styles/tokens.css`. Theme the brand once, there.

## ⛔ Three hard rules

1. **Don't touch the engine** — leave `src/deck/` (the engine + chrome: `Deck`, `Slide`, `Build`, `Reveal`, `DeckContext`, `useInView`, icons, `Annotator`) and `src/components/` untouched. Never edit `src/styles/base.css`.
2. **Author, don't reskin** — `src/App.tsx` is throwaway starter content. Delete its slides; author the real deck.
3. **Center what stands alone.** Ask every slide: *does it have a side visual* (`Split` media panel, image, `BrowserFrame`, chart beside text)?
   - **No side visual** (only text, or one structured block like `Comparison` / `Tabs` / `Timeline` / `Accordion` / `StatGrid`) → slide MUST be centered: use `<Slide center>`, `textAlign:'center'` on the heading, `marginInline:'auto'` on the block below it.
   - **Yes** → left-aligned/asymmetric allowed; the visual balances the text.
   A lone left-anchored block floating in empty space is the #1 alignment bug — never ship one.

---

## Step 0 — Ground the deck in the user's real input

Use the user's real topic, brand, document, facts. Never fabricate a placeholder company, logo, or quote. If a URL/brand is given, the theme comes from the brand — fetch the page, get real colors/font/logo, use the brand's known palette, then **STOP and ask**. Report the colors/fonts you used and where they came from.

---

## Step 1 — Run it in place (the repo is the app)

The repo root is already a complete Vite + React app — no scaffolding or copying. Layout:

```
package.json
vite.config.ts
tsconfig*.json
index.html
src/
  main.tsx
  App.tsx              ← THROWAWAY. Delete its slides; author the real deck.
  styles/
    tokens.css         (edit :root ONLY)
    base.css           ← don't edit base.css
  deck/
    Deck Slide Build Reveal DeckContext useInView icons Annotator
                       ← engine + UI. LOCKED.
  components/
    Cover BigNumber Contrast Chat Bento Split StatGrid Section
    Quote Pricing Steps Agenda Team CountUp TiltCard Marquee
    VisualDashboard Accordion Comparison Tabs Timeline CodeWindow
    BrowserFrame SpotlightCard Charts
```

`npm install && npm run dev` runs the deck at `/`. Verify dock / thumbnail rail / click-builds work, then delete the starter slides and author the real deck in `App.tsx`.

---

## Step 2 — Theme it (edit only the `:root` block)

All color, type, radius, depth, motion live in `src/styles/tokens.css` `:root`. **Change values, never variable names.** Nine ready-made theme families to pull from (dark product, editorial luxury, Swiss, dark technical, warm minimal, fintech, aurora glass, cinematic, paper editorial). One accent, used sparingly.

Dark vs light: set `html { color-scheme }` in base.css, pick `--bg`/`--fg` accordingly. Set fonts in `--font-head`/`--font-body` with `@import` at the top of `base.css`. Derive from the brand if given.

**Tab title + icon — always, unprompted.** Shared decks show the browser tab, so never leave `index.html` placeholders: set `<title>` to the deck's real title (e.g. "Acme — Series A") and swap the emoji in the favicon `<link>` to one that fits the topic. Do this for every deck without being asked.

---

## Step 3 — Author slides (each child of `<Deck>` is one slide)

Compose slides in `App.tsx`. Building blocks:

- **`<Slide>`** — one slide. `center` for statement/CTA; `full` for edge-to-edge; `nav="Label"`; `notes="…"` (editable in presenter overlay).
- **`<Cover>`** — opening slide: kicker → display title → subtitle cascade, full-bleed `image` under theme-correct scrim, optional `foot` line.
- **`<BigNumber>`** — ONE enormous accent figure (pass `<CountUp>`) + caption. Every deck gets one drama beat.
- **`<Contrast>`** — before/after, problem→solution: muted panel vs accent-lit panel with cross/check points.
- **`<Split>`** — text + edge-to-edge media (`flip` swaps sides). Media = `<img>`, color panel, `<BrowserFrame>`, or `<TiltCard><VisualX/></TiltCard>`.
- **`<Bento>`** — asymmetric tile grid; tiles take `c`/`r` spans + `variant`.
- **`<StatGrid>`** — responsive proof cards; pass `<CountUp>` as the stat `value`.
- **`<Section>`** — chapter divider: ghost number + accent glows. Full-bleed breather for decks without photography; use between parts.
- **`<Agenda>`** — numbered table-of-contents rows (strings, or `{title, hint}`).
- **`<Steps>`** — horizontal numbered process; connector draws in. Use for "how it works" instead of a bulleted list.
- **`<Pricing>`** — 2–4 tier cards; `highlight: true` crowns one with a badge.
- **`<Team>`** — people grid; photos via `img`, else auto-initials on accent.
- **`<Quote>`** — pull-quote slide with attribution (don't add quotation marks — the accent mark provides them).
- **`<Comparison>`** — us-vs-them feature matrix; one column highlighted in accent.
- **`<Table>`** — real data table: uppercase ruled header, right-aligned tabular numerals, optional `highlightCol`/`highlightRow`, `caption` source. Keep ≤5 columns / ≤7 rows — paged slides can't scroll.
- **`<Tabs>`** — tabbed content with a sliding accent pill.
- **`<Accordion>`** — expand/collapse panels (FAQ, feature detail).
- **`<Timeline>`** — vertical roadmap that draws its connector + milestones in.
- **`<CodeWindow>`** — macOS code window with line numbers + line highlight.
- **`<BrowserFrame>`** — browser chrome around full-bleed edge-to-edge content.
- **`<SpotlightCard>`** — cursor-follow glow card.
- **`<BarChart>` / `<LineChart>` / `<DonutChart>`** — draw-in CSS charts.
- **`<Chat>`** — ONLY if the product is genuinely conversational / AI interface. Never decoration for non-chat products.
- **`<Globe>`** — ONLY when the story is genuinely geographic. Markers must be REAL locations.

The workhorses are `<Slide>`, `<Cover>`, `<Split>`, `<Bento>`, `<StatGrid>` atoms; specialty layouts appear **at most once each**, when content calls for them. If you can't say in one sentence why a layout serves *this* deck, cut it.

**Compose like web, not like slideware**: full-bleed, layered; `Bento`/`Split` over centered rows of equal cards; oversized type with one accent word; vary rhythm so no two adjacent slides share shape; one idea per slide; open on cover, close on CTA.

**Click-builds** — `<Build at={N}>` reveals on the Nth click (space / Next), then moves to the next slide. Use for: punchline after setup, each step of a process, items appearing in turn.

**`<Reveal>`** — on-enter entrance (no click needed) on headlines/grids.

```tsx
<Slide center nav="The shift" notes="Pause, then reveal each point.">
  <h2 className="headline" style={{ marginInline: 'auto' }}>Three things changed.</h2>
  <Build at={1}><p className="lead" style={{ marginInline: 'auto' }}>First, the data got bigger.</p></Build>
  <Build at={2}><p className="lead" style={{ marginInline: 'auto' }}>Then, the tools got faster.</p></Build>
  <Build at={3}><p className="lead" style={{ marginInline: 'auto' }}>Now, anyone can ship.</p></Build>
</Slide>
```

---

## Step 4 — Responsive, not fixed (no clipping)

Each slide is a **full-viewport responsive layout**, not a fixed canvas — it reflows to any screen, so nothing scales-and-clips:

- **Fluid sizing.** Atoms use `clamp()`; use `%`, `vw`, `rem`, `max-width` containers — not fixed pixel widths that break on small screens.
- **Never hand-write a fixed column count** (`repeat(3, 1fr)` clips on phones). Use the **`.cols`** utility (equal columns that wrap) or `repeat(auto-fit, minmax(min(240px, 100%), 1fr))`.
- **One idea per slide**, sized to fill ~one screen with deliberate negative space.
- **Check narrow viewport** — `Bento`/`Split`/`Steps`/`Pricing`/`Contrast`/`Team` stack or compact themselves (built in); make sure headlines don't overflow and nothing needs scrolling — paged slides CANNOT scroll, overflow is truncation.
- **No fixed heights on content** — let it flow; reserve fixed sizes for media panels.

---

## Step 5 — Visuals & imagery

Visuals must fit the topic: data/SaaS → `.vframe` mock; brand/product/editorial/real-world → **generate images** into `public/`, one consistent style, used as `Split` media or full-bleed slide backgrounds under a gradient scrim (no text in images). A `Split` full-bleed image beats a floating card.

**Image-capable layouts** — when the deck calls for photography, take it directly (each puts a theme-correct scrim under text automatically).

**Motion** — ambient background (drifting spotlights + grain + vignette) and slide-change transitions are automatic. **One or two motion ideas per slide**, never a circus. All honors `prefers-reduced-motion`.

---

## Step 6 — Extend the system — invent new slides, components & visuals

The kit is a **floor, not ceiling.** The bundled components cover a lot — but author new ones when the topic calls for something nothing fits: a Gantt, device/phone mock, chat or kanban mock, map. Only the token *names* in `src/deck/` (engine + chrome) are off-limits to rewrite; **adding** components/visuals is encouraged.

Every new piece must: use `var(--…)` tokens only (no raw hex), compose like a web section, be responsive (work on mobile), animate with `Reveal`/`Build`/`useInView` + honor reduced-motion, use tabular figures, and add **no new dependencies** (plain React + CSS + SVG).

---

## Step 7 — Structure & writing

Pick the arc that fits the deck type (pitch, launch, brand, teaching, report) and let structure follow content. Open on cover, close on CTA. ~8–16 slides sized to the material. Headlines short, declarative, specific (sentence case); body 1–3 tight sentences; 1–3 word kickers; one idea per slide. Use the user's real numbers; never invent numbers for a real brand. Zero lorem, zero placeholder names. Add `notes` with talking points where useful.

---

## Pre-delivery checklist

- [ ] `src/deck/` untouched; step back through builds, fullscreen / sidebar (S) / grid view (G) work, annotation (A) tools persist per slide, presenter (P) opens a synced new tab, `H` hides UI, URL hash tracks slide.
- [ ] Deck is **authored, not reskinned** — topic, structure, copy, names are the user's; no starter leftovers.
- [ ] If a brand/URL was given, `--primary`, fonts, logo come from the brand.
- [ ] `index.html` has the deck's real `<title>` and a topic-matched favicon emoji — no placeholder left behind.
- [ ] Only the `:root` block was edited for theme; editing `--primary` recolors everything including the dock.
- [ ] Slides compose like web sections (full-bleed/asymmetric/bento/split), not centered card rows; visuals fit the topic.
- [ ] **Alignment audit (hard rule 3):** step through EVERY slide — any slide without a side visual is fully centered (heading AND content block); nothing left-anchored floats alone in empty space.
- [ ] **No showcase decks:** every specialty layout used meets its entry condition. Nothing included just because the kit has it.
- [ ] **Responsive:** looks right narrow + wide — sections stack, nothing clips or needs scrolling. Builds reveal in intended order.
- [ ] Motion restrained; reduced-motion respected.
- [ ] `npm install && npm run dev` runs with no console errors; `npm run build` passes.
