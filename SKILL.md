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

Create publication-grade, human-designed PowerPoint (`.pptx`) presentations tailored to the exact subject, audience level, and language (including full Azerbaijani diacritics: `Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç`).

---

## Mandatory 2-Phase Token-Efficient Workflow

### Phase 1: Compact Design Brief (Wait for Confirmation)
When the user specifies a topic, discipline, and audience level, respond first with a **concise Design Brief**:
1. **Sanzo Wada Palette**: Pick 1 combination ID from [sanzo-wada-palettes.json](./references/sanzo-wada-palettes.json) ensuring visual unity across all slides.
2. **Typography**: Pick clean, eye-friendly Azerbaijani-safe fonts from [typography-matrix.md](./references/typography-matrix.md) matched to the subject.
3. **Custom Background & Visual Direction**: Describe the unified background/surface direction and verified real/scientific visuals to source.
4. **Slide-by-Slide Plan**: List every content slide, a dedicated **References** slide near the end, and a mandatory **Thank You / Təşəkkürlər** closing slide—specifying consistent grid placement, automatic Slide Show-safe animations, and intentional `Morph`/`Merge Shapes` usage.

Wait for user confirmation before researching full content or generating `.pptx`.

### Phase 2: Sourcing, Native COM Build, and Visual QA
1. **Targeted Research & Asset Sourcing** (follow [research-and-assets.md](./references/research-and-assets.md)):
   - Prefer `search_web` and direct open APIs first; use `browser` only when visual curation or JS-rendered sites require it.
   - Verify all facts and numbers against trusted academic/scientific databases (`OpenAlex`, `arXiv`, `PubMed`, `Britannica`).
2. **Write Compact `deck_spec.json`**:
   - Use 16:9 coordinates (`960 x 540` pt) with a **consistent, predictable slide-to-slide placement grid** for titles, body text, and imagery.
3. **Build & Export PNGs via Reusable COM Engine**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "C:\Users\tehsil\.gemini\config\skills\bespoke-pptx-studio\scripts\build_deck.ps1" -SpecPath "<path\to\deck_spec.json>" -OutputPath "<path\to\output.pptx>" -ExportPngDir "<path\to\qa_pngs>"
   ```
   *(If PowerPoint COM is blocked inside the terminal sandbox, run `build_deck.ps1` outside the sandbox while keeping research/prep steps inside the sandbox.)*
4. **Visual & Slide Show QA (Mandatory)**:
   - Open ALL exported `Slide*.PNG` files using `read_file` and verify:
     - Zero `<14 pt` micro-text, zero AI-sign kickers, and zero inline image footnotes.
     - Zero clashing high-contrast boxes or pasted rectangular borders around images.
     - Consistent title, text, and image alignment across content slides.
     - Zero text clipping/overflow and crisp Azerbaijani diacritics (`ə, ğ, ı, ö, ş, ü, ç`).
   - Verify animation timelines contain no missing targets, empty frames, or multi-click stalls so Slide Show mode advances smoothly.

---

## Anti-AI-Slop & Human-Design Rules

1. **Ban "AI-Sign" Micro-Text & Kickers**: Never place tiny (`<14 pt`) uppercase labels, section tags (`01 • SECTION`), or small header/footer metadata strips on slides.
2. **Clean Natural Punctuation**: Never put synthetic symbols (`~`, `•`, `|`, `—` em-dash clutter, or `01 •` prefixes) in headings or titles. Use natural prose and clean typography.
3. **References Only at the End**: Never clutter content slides with tiny captions or source footnotes under images. Consolidate all citations cleanly on a dedicated **References / İstifadə Olunmuş Mənbələr** slide near the end.
4. **Mandatory Closing Slide**: Always conclude the deck with a clean, dedicated closing slide (e.g., **"Diqqətinizə görə təşəkkürlər!"** / **"Təşəkkürlər"**).
5. **Visual Unity (Zero Clashing Blocks)**: Never drop jarring high-contrast boxes onto mismatched backgrounds (such as harsh dark boxes on light slides or bordered white boxes on dark slides). Integrate diagrams seamlessly onto a cohesive canvas.
6. **Never Fall Back to Blank Solid White (`#ffffff`) Backgrounds**: Fixing diagram integration or visual unity MUST NEVER remove the custom-searched background. Always preserve an authentic, custom-searched background (e.g., fine cotton/watercolor paper grain, linen, parchment, or domain-matched atmospheric photography) paired with a calibrated Sanzo Wada tint overlay across all slides so the deck stays tactile, unified, and unmistakably bespoke.
7. **Consistent Slide-to-Slide Grid**: Maintain steady, predictable coordinates for slide titles, text columns, and images from slide to slide instead of jumping layouts arbitrarily left and right.
8. **Ban Template Clichés & Hallucinations**: Zero purple-blue neon gradients, zero 3-card icon grids, zero Corporate-Memphis vectors, and zero fabricated statistics. Build natively in PowerPoint COM.

---

## PowerPoint Functionality Standards

- **Slide Show-Safe Dynamic Animations**: Choreograph smooth, readable sequences (`fade`, `fly_in`, `wipe`, `zoom`) using `with_previous` and `after_previous` timings (0.35–0.55 s) so slides reveal naturally upon entry without trapping the presenter in awkward multi-click chains or broken Slide Show pacing.
- **Morph Transitions (`!!` Prefix)**: When consecutive slides share a continuous element, prefix its name with `!!` and set `transition.type = "morph"` (`build_deck.ps1` falls back cleanly to smooth fade `3849` on Office 2016 RTM builds that predate Morph, and runs native Morph on Office 2019/365).
- **Per-Slide `Merge Shapes` Assessment**: Apply `intersect`, `subtract`, `union`, `combine`, or `fragment` only when a clean photo mask or composite shape strengthens editorial quality without introducing visual clutter.
