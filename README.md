# Bespoke PPTX Studio

[English](#english) | [Azərbaycanca](#az%C9%99rbaycanca)

---

## English

An Antigravity agent skill for building **bespoke, human-designed PowerPoint (`.pptx`) presentations** powered by:

- **Authentic Sanzo Wada Color Harmonies**: 16 verified 2-, 3-, and 4-color palettes extracted from *A Dictionary of Color Combinations* (`sanzo-wada.dmbk.io`) applied with calm visual unity (no clashing boxes).
- **Custom-Searched Real Backgrounds & Imagery**: Editorial textures, archival plates, and verified scientific diagrams integrated cleanly with `fit: "contain"` or `fit: "cover"` support.
- **Trained Anti-AI-Slop Standards**: Zero tiny `<14 pt` micro-text or AI-sign kickers, clean natural punctuation, consistent slide-to-slide layout grid, consolidated **References** at the end, and a mandatory **Thank You** closing slide.
- **Eye-Friendly Azerbaijani-Safe Typography**: Discipline-specific font pairings tested for full Azerbaijani diacritics (`Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç`).
- **Native PowerPoint COM Engine (`scripts/build_deck.ps1`)**: Builds `.pptx` directly via desktop Microsoft PowerPoint COM—supporting native **Morph (`!!`)**, **Merge Shapes** (`Union`, `Intersect`, `Subtract`, `Combine`, `Fragment`), Slide Show-safe **Entrance + Emphasis animations**, and **1920×1080 PNG export** for visual QA.

### Installation

```powershell
git clone https://github.com/nsc0806/bespoke-pptx-studio "$env:USERPROFILE\.gemini\config\skills\bespoke-pptx-studio"
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

Standart AI şablonlarından tam uzaqlaşan, insan gözünə təbii və yüksək keyfiyyətli **PowerPoint (`.pptx`)** təqdimatları hazırlayan Antigravity skill-i:

- **Sanzo Wada Rəng Harmoniyası**: `sanzo-wada.dmbk.io` mənbəyindən çıxarılmış 16 orijinal 2, 3 və 4 rəngli palitra — təzadlı qutu yığını olmadan vahid vizual bütövlük.
- **Xüsusi Axtarılmış Real Fonlar və Elmi Vizuallar**: Mövzuya uyğun teksturalar, real fotolar və təmiz `fit: "contain"` / `fit: "cover"` diaqram yerləşdirməsi.
- **Təkmilləşdirilmiş Anti-AI-Slop Qaydaları**: Xırda `<14 pt` "AI-sign" yazıları və başlıq durğu işarələri qadağandır; slaydlarda ardıcıl və sabit yerləşmə sistemi, sonda ayrıca **Mənbələr** və **Təşəkkürlər** slaydı məcburidir.
- **Azərbaycan Əlifbası Dəstəkli Rahat Tipoqrafiya**: `Ə, ə, Ğ, ğ, İ, ı, Ö, ö, Ş, ş, Ü, ü, Ç, ç` simvollarını pozmayan, gözü yormayan şrift cütləri.
- **Native PowerPoint COM Mühərriki (`scripts/build_deck.ps1`)**: Birbaşa desktop PowerPoint vasitəsilə **Morph (`!!`)**, **Merge Shapes**, Slide Show rejimində problemsiz işləyən **dinamik animasiyalar** və **PNG vizual QA** ixracı yaradır.
