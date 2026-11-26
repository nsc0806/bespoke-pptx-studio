# Token-Efficient Research & Asset Sourcing Guide

Follow strict **database targeting** and **tool priority** to maximize theoretical accuracy and visual authenticity while minimizing wasted tokens.

---

## 1. Tool Priority (Speed & Token Budget)

1. **`search_web` & Direct Open APIs (Default First Choice)**:
   - Use `search_web` or lightweight PowerShell `Invoke-RestMethod` queries against open JSON/REST APIs (`OpenAlex`, `arXiv`, `Wikimedia Commons`, `Unsplash`/`Pexels`, `Met Museum`) to gather verified facts, dates, metrics, and direct high-resolution image URLs in 1–2 calls.
2. **`browser` Subagent (Targeted Escalation Only)**:
   - Use ONLY when:
     - Visually curating among multiple photo candidates or checking exact color/mood alignment on a JS-rendered site.
     - Target pages block simple HTTP requests or require interactive navigation.
   - Always pass a narrow, concrete completion criterion so the browser subagent finishes in minimal steps.

---

## 2. Custom Background & Imagery Sources (No Generic AI Art)

| Visual Need | Preferred Sources | Query Strategy |
|---|---|---|
| **Atmosphere, Architecture, Landscapes, Workspaces** | Unsplash, Pexels, Wikimedia Commons | Search specific mood + material + lighting (e.g., `"brutalist concrete shadow morning light"`, `"minimalist botanical herbarium paper"`). |
| **Fine Art, History, Manuscripts, Cultural Artifacts** | Metropolitan Museum Open Access, Smithsonian Open Access, Wikimedia Commons (`File:` high-res) | Search exact historical period, artist, museum plate, or archival keyword; prefer `>= 1920x1080` landscape assets for 16:9 full-bleed slides. |
| **Tactile Backgrounds (Paper, Linen, Stone, Parchment)** | Wikimedia Commons, Unsplash, Pexels | Search macro surface textures (`"washi paper texture warm"`, `"linen canvas macro"`, `"terracotta plaster texture"`), then pair with Sanzo Wada overlay transparency. |
| **Scientific / Technical Figures** | Wikimedia Commons, NASA Images, OpenStax, arXiv / PubMed Central open figures | Prefer clean schematic plates or real observatory/lab photography over cartoonish vector icons. |

### Image Validation Checklist Before Embedding
- Minimum width `1600 px` (ideal `>= 1920 px`) for full-bleed backgrounds; minimum `800 px` for split-panel photos.
- Verify downloaded file size (`> 20 KB`) and valid image format (`.jpg`, `.jpeg`, `.png`) before passing to `build_deck.ps1`.
- Match photo orientation (landscape vs. portrait) to its container so faces or focal subjects are never distorted.

---

## 3. Theoretical Data & Fact Verification Sources

- **Peer-Reviewed & Academic**: `OpenAlex` (`api.openalex.org`), `arXiv`, `PubMed`, `Semantic Scholar`
- **Encyclopedic & Historical Reference**: `Encyclopaedia Britannica`, `UNESCO`, national museum archives, official statistical agencies (`World Bank`, `Eurostat`, `Stat.gov.az`)
- **Rule**: Cite exact figures, years, and authors in subtle slide footnotes (`10.5–12 pt`) when presenting empirical claims.
