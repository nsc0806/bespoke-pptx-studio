# Bespoke PPTX Studio

[English](#english) | [Azərbaycanca](#az%C9%99rbaycanca)

---

## English

An Antigravity agent skill for building **bespoke, non-generic PowerPoint (`.pptx`) presentations** powered by:

- **Authentic Sanzo Wada Color Harmonies**: 16 verified 2-, 3-, and 4-color palettes extracted from *A Dictionary of Color Combinations* (`sanzo-wada.dmbk.io`).
- **Custom-Searched Real Backgrounds & Imagery**: Editorial textures, archival museum plates, and real photography paired with Sanzo Wada color overlays instead of generic AI clipart or purple gradients.
- **Anti-AI-Slop Rules (Reddit Community-Backed)**: Breaks uniform slide rhythm, eliminates 3-card icon clichés, enforces one dominant narrative hook per slide, and prohibits hallucinated metrics.
- **Subject-Adaptive Azerbaijani-Safe Typography**: Discipline-specific font pairings tested for full Azerbaijani diacritics (`Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç`).
- **Native PowerPoint COM Engine (`scripts/build_deck.ps1`)**: Builds `.pptx` directly via desktop Microsoft PowerPoint COM—supporting native **Morph (`!!`)**, **Merge Shapes** (`Union`, `Intersect`, `Subtract`, `Combine`, `Fragment`), dynamic step-by-step **Entrance + Emphasis animations**, and **1920×1080 PNG export** for visual QA.

### Installation

Clone or copy into your Antigravity global skills directory:

```powershell
git clone <your-repo-url> "$env:USERPROFILE\.gemini\config\skills\bespoke-pptx-studio"
```

*(Requires Windows + Microsoft PowerPoint desktop installed.)*

### Quick Test

```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\build_deck.ps1" `
  -SpecPath ".\deck_spec.json" `
  -OutputPath ".\output.pptx" `
  -ExportPngDir ".\qa_pngs"
```

---

## Azərbaycanca

Standart AI şablonlarından tam uzaqlaşan, fərdi və yüksək keyfiyyətli **PowerPoint (`.pptx`)** təqdimatları hazırlayan Antigravity skill-i:

- **Sanzo Wada Rəng Harmoniyası**: `sanzo-wada.dmbk.io` mənbəyindən çıxarılmış 16 orijinal 2, 3 və 4 rəngli palitra.
- **Xüsusi Axtarılmış Real Fonlar və Vizuallar**: Mövzuya uyğun kağız/kətan teksturaları, muzey arxivləri və real fotolar üzərində Sanzo Wada overlay harmoniyası.
- **Reddit Tənqidlərinə Əsaslanan Anti-AI-Slop Qaydaları**: Təkrarlanan eyni slayd ritmi, 3-kartlı şablonlar, boş vektorlar və uydurma rəqəmlər qadağandır.
- **Azərbaycan Əlifbası Dəstəkli Tipoqrafiya**: `Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç` simvollarını pozmayan fənnə-adaptiv şrift cütləri.
- **Native PowerPoint COM Mühərriki (`scripts/build_deck.ps1`)**: Birbaşa desktop PowerPoint vasitəsilə **Morph (`!!`)**, **Merge Shapes**, addım-addım **dinamik animasiyalar** və **PNG vizual QA** ixracı yaradır.
