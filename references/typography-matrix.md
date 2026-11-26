# Subject-Adaptive Azerbaijani-Safe Typography Matrix

Every chosen font **must** render full Azerbaijani Latin Extended characters cleanly (`Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç`) without glyph fallback mismatches.

---

## 1. Curated Discipline Pairings

| Discipline / Tone | Display / Heading Font | Body / Data Font | Optional Mono / Label | Character & Rationale |
|---|---|---|---|---|
| **History, Literature, Humanities** | `Georgia` (or `Playfair Display`) | `Aptos` / `Calibri` (or `Lora`) | `Consolas` | Editorial serif authority paired with clean legible body text. |
| **Academic Science & Medicine** | `Cambria` (or `Source Serif 4`) | `Segoe UI` (or `IBM Plex Sans`) | `JetBrains Mono` | Crisp scientific hierarchy; distinct numerals and clear subscript readability. |
| **Engineering, IT & Architecture** | `Bahnschrift` (or `Space Grotesk`*) | `Segoe UI` (or `Inter`) | `JetBrains Mono` | Precision industrial geometry; technical clarity for specs and diagrams. |
| **Business, Strategy & Economics** | `Sitka Display` / `Georgia` | `Aptos` / `Calibri` (or `Plus Jakarta Sans`) | `Consolas` | Executive contrast between refined titles and scannable body copy. |
| **Art, Culture, Design & Museums** | `Palatino Linotype` (or `Fraunces`*) | `Segoe UI` (or `Inter`) | `Consolas` | Gallery-grade proportions that frame archival imagery without overpowering it. |
| **Botanical, Ecology & Geography** | `Constantia` (or `Lora`) | `Calibri` (or `IBM Plex Sans`) | `Consolas` | Organic serif terminals balanced with airy sans-serif captions. |

*\*Note: Before using optional Google Fonts (`Playfair Display`, `Lora`, `Inter`, `Plus Jakarta Sans`, `IBM Plex Sans`, `Source Serif 4`, `JetBrains Mono`), verify they are installed in `C:\Windows\Fonts` or `%LOCALAPPDATA%\Microsoft\Windows\Fonts`, and confirm `Ə/ə` glyph coverage. Windows-native fonts (`Georgia`, `Cambria`, `Palatino Linotype`, `Sitka Display`, `Constantia`, `Bahnschrift`, `Segoe UI`, `Aptos`, `Calibri`, `Consolas`) are pre-verified fallback-safe on Windows Office 16+.*

---

## 2. Scale & Hierarchy Rules (16:9 Canvas: `960 x 540` pt)

- **Hero Cover Title**: `40 – 54 pt` (Bold, tight line spacing `1.05 – 1.12`)
- **Slide Title**: `28 – 36 pt` (Bold or Semibold, never wrap beyond 2 lines)
- **Section Kicker / Eyebrow**: `11 – 13 pt` (Uppercase or tracked label in accent color)
- **Primary Body Copy**: `16 – 20 pt` (Line spacing `1.2 – 1.35`; never drop below `15 pt` for prose)
- **Key Metric / Big Number**: `48 – 72 pt` (Single focal anchor per data slide)
- **Captions & Source Citations**: `10.5 – 12 pt` (Muted palette color, crisp contrast)

---

## 3. Text-to-Visual Balance & Density Calibration

- **Audience Calibration**:
  - *School / Public / Pitch*: 15–35 words per slide; 55%+ visual/whitespace.
  - *University Lecture / Academic*: 40–80 words max per slide (or structured comparison/diagram); 35–45% visual/whitespace.
  - *Technical / Expert Review*: Structured tables, annotated charts, or split code/diagram layouts—never walls of unformatted prose.
- **Line Width Rule**: Keep prose text boxes between `45–65` characters per line (`340–480 pt` wide on 16:9) for effortless reading.
- **Contrast Floor**: Text over custom background imagery must sit on either a Sanzo Wada semi-transparent overlay (`25–45%` opacity) or a dedicated Sanzo Wada surface panel with clear `>= 4.5:1` contrast.
