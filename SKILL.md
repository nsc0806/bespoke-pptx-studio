---
name: bespoke-pptx-studio
description: >-
  Creates bespoke, non-generic PowerPoint (.pptx) presentations using Sanzo Wada
  color combinations, custom-searched authentic backgrounds and imagery,
  subject-adaptive Azerbaijani-safe typography, native PowerPoint COM Morph,
  Merge Shapes, and dynamic step-by-step animations with minimal token usage.
  Use whenever the user asks to design, plan, or generate a PowerPoint deck.
---

# Bespoke PPTX Studio

Create publication-grade, non-generic PowerPoint (`.pptx`) presentations tailored to the exact subject, audience level, and language (including full Azerbaijani diacritics: `Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç`).

---

## Mandatory 2-Phase Token-Efficient Workflow

### Phase 1: Compact Design Brief (Wait for Confirmation)
When the user specifies a topic, discipline, and audience level, respond first with a **concise Design Brief**:
1. **Sanzo Wada Palette**: Pick 1 combination ID from [sanzo-wada-palettes.json](./references/sanzo-wada-palettes.json) matching the subject mood and define exact slide roles (`background`, `surface`, `text_primary`, `text_muted`, `accent`, `overlay`).
2. **Typography Pair**: Pick 1 Azerbaijani-safe pairing from [typography-matrix.md](./references/typography-matrix.md) matched to the discipline.
3. **Custom Background & Visual Direction**: Describe the authentic backgrounds/textures/archival or photographic visuals to search (never generic defaults).
4. **Slide-by-Slide Plan**: List each slide with:
   - Narrative hook (1 central takeaway) and layout rhythm (asymmetric split, editorial full-bleed, sparse hero statement, comparative table, etc.)
   - **Dynamic Animation Choreography**: step-by-step Entrance + Emphasis sequence (`on_click`, `with_previous`, `after_previous`)
   - **Transition / Morph (`!!`) & Merge Shapes Evaluation**: explicitly decide per slide whether `Morph` or `Merge Shapes` adds real functional/editorial value (use intentionally, never everywhere).

Wait for user confirmation before researching full content or generating `.pptx`.

### Phase 2: Sourcing, Native COM Build, and Visual QA
1. **Targeted Research & Asset Sourcing** (follow [research-and-assets.md](./references/research-and-assets.md)):
   - Prefer `search_web` and direct open APIs first for speed and token economy.
   - Use the `browser` subagent only when visual curation or JS-rendered pages require it.
   - Download verified high-resolution backgrounds and context-synchronized imagery locally into a scratch/assets folder; verify theoretical facts and numbers against trusted databases (`OpenAlex`, `arXiv`, `PubMed`, `Britannica`).
2. **Write Compact `deck_spec.json`**:
   - Define slides declaratively using 16:9 coordinates (`960 x 540` pt).
   - Keep text depth calibrated to the requested audience level and protect text-to-image balance.
3. **Build & Export PNGs via Reusable COM Engine**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "C:\Users\tehsil\.gemini\config\skills\bespoke-pptx-studio\scripts\build_deck.ps1" -SpecPath "<path\to\deck_spec.json>" -OutputPath "<path\to\output.pptx>" -ExportPngDir "<path\to\qa_pngs>"
   ```
   *(If PowerPoint COM is blocked inside the terminal sandbox, run `build_deck.ps1` outside the sandbox while keeping research/prep steps inside the sandbox.)*
4. **Visual QA (Mandatory)**:
   - Open exported `Slide*.PNG` files using `read_file` and verify:
     - Zero text clipping, box overflow, or overlapping elements.
     - High contrast between text and custom backgrounds (adjust Sanzo Wada overlay transparency if needed).
     - Clean rendering of Azerbaijani characters (`ə, ğ, ı, ö, ş, ü, ç`).
     - Balanced whitespace, text-to-image ratio, and distinct per-slide rhythm.
   - Fix any defect in `deck_spec.json`, rebuild, and re-verify before delivering the `.pptx`.

---

## Anti-AI-Slop Rules (Reddit Community Consensus)

Avoid the exact failure modes presentation professionals criticize in AI slide tools (`r/powerpoint`, `r/ClaudeAI`, `r/startups`):
1. **Break Uniform Rhythm**: Never repeat the same `Title + Subtitle + 3 bullets` cadence across consecutive slides. Alternate sparse hero frames (3–7 words), full-bleed editorial spreads, asymmetric 65/35 splits, and focused data views.
2. **Ban Template Clichés**: Zero purple-blue neon gradients, zero rounded 3-card feature grids with floating icons, and zero purposeless decorative blobs/circles.
3. **One Narrative Hook per Slide**: Never equalize all text to "look full." One dominant claim, number, or visual anchors each slide; supporting details stay subordinate.
4. **Authentic Visuals Only**: Zero Corporate-Memphis vectors or cheesy 3D clipart. Every background and photo must come from searched real photography, textures, archival plates, or domain-accurate diagrams.
5. **Native `.pptx` Integrity**: Never export brittle HTML-to-PPTX wrappers. Build natively through PowerPoint COM so every shape, mask, transition, and animation remains 100% editable.
6. **Zero Hallucinated Data**: Never invent statistics, dates, or citations; source and verify every claim.

---

## PowerPoint Functionality Standards

- **Dynamic & Rich Animations**: Choreograph step-by-step reveals (`fly_in`, `wipe`, `fade`, `zoom`) paired with purposeful emphasis (`pulse`, `grow_shrink`, `teeter`) on key numbers, terms, or focal elements. Sequence supporting beats with `with_previous` / `after_previous` so slides feel alive rather than cluttered with single-line clicks.
- **Morph Transitions (`!!` Prefix)**: For spatial continuity across consecutive slides (expanding a detail, shifting a timeline, reframing a visual), give matching objects the exact same name prefixed with `!!` (e.g., `!!hero_frame`) and set `transition.type = "morph"` (`build_deck.ps1` automatically falls back to smooth fade `3849` on Office 2016 RTM builds that predate Morph, while running native Morph on Office 2019/365).
- **Per-Slide `Merge Shapes` Assessment**: Evaluate every slide individually—apply `intersect`, `subtract`, `union`, `combine`, or `fragment` when crafting custom editorial image masks, split cutouts, or bespoke composite geometry, and skip it when clean typography or full-bleed photography is stronger.
