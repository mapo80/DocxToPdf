# DocxToPdf.Sdk

**Motore di conversione DOCX → PDF in C# puro** basato su **SkiaSharp** (rendering PDF) e **HarfBuzz** (text shaping avanzato).

## Stato del progetto

✅ **Fase 1 completata:** "Hello, World" - Kernel PDF + Text Shaping
✅ **Fase 2 completata:** Conversione DOCX → PDF con layout automatico

Il progetto implementa un convertitore DOCX → PDF completamente funzionante, con supporto per testo, margini, paginazione automatica e font fallback.

## Architettura

### Moduli implementati

#### 1. `DocxToPdf.Sdk.Pdf` - Backend PDF
- **`PdfDocumentBuilder`**: Wrapper sopra `SKDocument.CreatePdf` con gestione ciclo vita (`BeginPage`/`EndPage`/`Close`)
- **`PdfPage`**: Rappresenta una pagina PDF con coordinate in punti tipografici (pt)
- **`PaperSize`**: Dimensioni standard carta (A4, Letter, Legal) in pt
- **`PdfMetadata`**: Metadati XMP del documento (Title, Author, etc.)

#### 2. `DocxToPdf.Sdk.Text` - Text Shaping
- **`TextRenderer`**: Rendering testo con HarfBuzz shaping (legature, diacritici, script complessi)
- **`FontManager`**: Gestione font di sistema con caching e font fallback automatico

#### 3. `DocxToPdf.Sdk.Units` - Sistema coordinate
- **`UnitConverter`**: Conversioni tra pt, mm, cm, inch, DXA e twips (unità WordprocessingML)
- **`Margins`**: Rappresentazione margini pagina con conversione da/verso DXA

#### 4. `DocxToPdf.Sdk.Docx` - Parsing DOCX (Nuovo!)
- **`DocxDocument`**: Wrapper per `WordprocessingDocument` con accesso semplificato
- **`DocxSection`**: Estrazione page size, margini e orientamento da `w:sectPr`
- **`DocxParagraph`** / **`DocxRun`**: Parser per testo e proprietà di formattazione

#### 5. `DocxToPdf.Sdk.Layout` - Layout Engine (Nuovo!)
- **`TextLayoutEngine`**: Layout automatico con word wrapping greedy
- **`DocxToPdfConverter`**: Convertitore completo DOCX → PDF con paginazione

## Caratteristiche tecniche

### Stack tecnologico
- **.NET 9.0** (C# 13)
- **SkiaSharp 3.119.1** - Rendering PDF nativo con backend cross-platform
- **SkiaSharp.HarfBuzz 3.119.1** - Text shaping di qualità tipografica
- **DocumentFormat.OpenXml 3.3.0** - Lettura documenti DOCX

### Unità di misura
- **Punto tipografico (pt)**: 1 pt = 1/72 inch (unità base del sistema)
- **Twips (DOCX page size)**: 1 pt = 20 twips, 1440 twips = 1 inch
- **DXA (DOCX margins)**: equivalente a twips, 1 pt = 20 dxa
- **A4**: 595.276 × 841.890 pt (210 × 297 mm)

### Text rendering
- **HarfBuzz shaping**: gestisce automaticamente legature (fi, fl, ffi, ffl), kerning, diacritici
- **Font fallback automatico**: usa `SKFontManager.MatchCharacter` per trovare font che supportano emoji, CJK e simboli speciali
- **Font di sistema**: usa i font disponibili su macOS/Windows/Linux
- **Antialiasing**: ottimizzato per output PDF (grayscale, non subpixel LCD)
- **Testo vettoriale**: tutto il testo è renderizzato come vettori (selezionabile, non raster)

## Demo e utilizzo

### 1. Demo "Hello, World"

Genera un PDF A4 di test con caratteri speciali, legature ed emoji:

```bash
cd DocxToPdf.Demo
dotnet run
```

Output: `hello.pdf` (circa 856 KB, PDF 1.4, 1 pagina)

### 2. Conversione DOCX → PDF

Converte un documento DOCX in PDF preservando margini, layout e formattazione base:

```bash
cd DocxToPdf.Demo
dotnet run samples/lorem.docx samples/lorem.pdf
```

Esempio con file di sistema:

```bash
dotnet run /path/to/documento.docx /path/to/output.pdf
```

**Formattazione supportata:**
- ✅ Dimensione pagina e orientamento (`w:pgSz`)
- ✅ Margini del documento (`w:pgMar`)
- ✅ Font famiglia, grassetto, corsivo (`w:rFonts`, `w:b`, `w:i`)
- ✅ Cascata stili Word (doc defaults → basedOn → pStyle/rStyle → direct formatting)
- ✅ Theme font/colors + `clrSchemeMapping`, tint/shade e colori hyperlink
- ✅ Dimensione font (`w:sz` in half-points)
- ✅ Paragrafo: spacing before/after, indentazioni (first-line/hanging) e allineamento base
- ✅ Liste numerate/puntate (multi-level `w:numPr`, `w:abstractNum` con restart/continuation)
- ✅ Wrapping automatico con paginazione
- ✅ Emoji e caratteri Unicode con font fallback

### 3. Sample "render" e temi

Il comando `render` consente di testare rapidamente i documenti di esempio focalizzati su stili e temi.

```bash
cd DocxToPdf.Demo
dotnet run render samples/styles-theme.docx -o out/styles-theme.pdf
dotnet run render samples/styles-theme-alt.docx -o out/styles-theme-alt.pdf
```

- `styles-theme.docx`: tema Office di default (Calibri/Cambria, accenti blu/arancio)
- `styles-theme-alt.docx`: tema personalizzato (Times/Arial, palette verde/corallo) per verificare il mapping major/minor + colori accent.

### 4. Sample numbering

```bash
cd DocxToPdf.Demo
dotnet run render samples/numbering-multilevel.docx -o out/numbering-multilevel.pdf
dotnet run render samples/numbering-multilevel.docx -o out/numbering-multilevel.pdf --log-numbering
```

Questo file mostra livelli 0‑2, bullet Wingdings, restart (`numId` diverso) e continuazioni sullo stesso `numId`.
L'opzione `--log-numbering` stampa in console il tracciamento `numId/ilvl` risolto.

## API pubblica

### Esempio 1: Conversione DOCX → PDF (High-level)

```csharp
using DocxToPdf.Sdk;

var converter = new DocxToPdfConverter();
converter.Convert("input.docx", "output.pdf");
```

### Esempio 2: Rendering PDF manuale (Low-level)

```csharp
using DocxToPdf.Sdk.Pdf;
using DocxToPdf.Sdk.Text;
using DocxToPdf.Sdk.Units;
using SkiaSharp;

// Metadati
var metadata = new PdfMetadata
{
    Title = "Il mio documento",
    Author = "Nome Autore",
    Creator = "DocxToPdf.Sdk"
};

// Crea documento
using var builder = PdfDocumentBuilder.Create("output.pdf", metadata);

// Inizia pagina A4
var page = builder.BeginPage(PaperSize.A4);
var margins = Margins.Default; // 72 pt = 1 inch

// Rendering testo con HarfBuzz + font fallback
var renderer = new TextRenderer();
var typeface = FontManager.Instance.GetDefaultTypeface();

renderer.DrawShapedTextWithFallback(
    page.Canvas,
    "Hello, world! 👋🌍",
    x: margins.Left,
    y: margins.Top + 48f, // baseline
    typeface,
    sizePt: 48f,
    SKColors.Black
);

// Termina pagina e chiudi
builder.EndPage();
builder.Close();
```

### Esempio 3: Lettura DOCX

```csharp
using DocxToPdf.Sdk.Docx;

using var docx = DocxDocument.Open("documento.docx");

// Leggi sezione (pagina + margini)
var section = docx.GetSection();
Console.WriteLine($"Pagina: {section.PageSize.WidthPt} × {section.PageSize.HeightPt} pt");
Console.WriteLine($"Margini: T:{section.Margins.Top} R:{section.Margins.Right} B:{section.Margins.Bottom} L:{section.Margins.Left}");

// Leggi paragrafi
foreach (var paragraph in docx.GetParagraphs())
{
    foreach (var run in paragraph.Runs)
    {
        Console.WriteLine($"Testo: '{run.Text}' (Font: {run.Formatting.FontFamily}, Size: {run.Formatting.FontSizePt}pt, Bold: {run.Formatting.Bold})");
    }
}
```

### Conversioni unità

```csharp
// Margini Word (DXA) → pt
var margins = Margins.FromDxa(
    top: 1440,    // = 72 pt = 1 inch
    right: 1440,
    bottom: 1440,
    left: 1440
);

// Conversioni manuali
float ptFromMm = UnitConverter.MmToPoints(25.4f); // = 72 pt
float ptFromDxa = UnitConverter.DxaToPoints(1440); // = 72 pt
```

## Struttura del progetto

```
DocxToPdf/
├── DocxToPdf.sln
├── DocxToPdf.Sdk/                      # Libreria principale (~1200 LOC)
│   ├── DocxToPdfConverter.cs           # Convertitore high-level
│   ├── Pdf/                            # Backend PDF (SkiaSharp)
│   │   ├── PdfDocumentBuilder.cs
│   │   ├── PdfPage.cs
│   │   ├── PaperSize.cs
│   │   └── PdfMetadata.cs
│   ├── Text/                           # Text shaping (HarfBuzz)
│   │   ├── TextRenderer.cs             # + font fallback automatico
│   │   └── FontManager.cs
│   ├── Units/                          # Conversioni unità
│   │   ├── UnitConverter.cs            # pt/mm/cm/inch/twips/dxa
│   │   └── Margins.cs
│   ├── Docx/                           # Parser DOCX
│   │   ├── DocxDocument.cs
│   │   ├── DocxSection.cs              # Page size + margini
│   │   ├── DocxParagraph.cs
│   │   └── DocxRun.cs                  # Testo + formattazione
│   └── Layout/                         # Layout engine
│       └── TextLayoutEngine.cs         # Wrapping greedy
├── DocxToPdf.Demo/                     # Applicazione demo
│   └── Program.cs                      # Hello World + converter
├── samples/                            # File di test
│   ├── lorem.docx                      # DOCX lorem ipsum
│   ├── styles-theme.docx               # Heading/Normal con tema Office
│   ├── styles-theme-alt.docx           # Stesso documento con tema alternativo
│   └── create_test_docx.py             # Script generatore lorem
└── README.md
```

## Roadmap

### ✅ Fase 1 - Kernel PDF & Text Shaping (Completata)
- [x] Backend PDF con SkiaSharp (`SKDocument.CreatePdf`)
- [x] Text shaping con HarfBuzz (`SKShaper`)
- [x] Sistema coordinate e unità (pt, DXA, mm, twips)
- [x] **Font fallback automatico** con `SKFontManager.MatchCharacter` (emoji, CJK, simboli)
- [x] Demo "Hello, World" funzionante
- [x] Testo vettoriale selezionabile

### ✅ Fase 2 - Integrazione DOCX (Completata)
- [x] Aggiungere **Open XML SDK** per lettura DOCX
- [x] Parser minimale: Body → Paragraph → Run → Text
- [x] Lettura page size e orientamento (`w:pgSz`)
- [x] Applicazione margini da `w:sectPr/w:pgMar` (conversione twips/DXA → pt)
- [x] Font mapping: famiglia, bold, italic, dimensione (`w:rFonts`, `w:b`, `w:i`, `w:sz`)
- [x] Layout engine con wrapping greedy (word breaking)
- [x] Paginazione automatica
- [x] Demo: conversione DOCX → PDF completa

### 🔄 Fase 3 - Miglioramenti tipografici (Prossima)
- [ ] Tipografia avanzata: **UAX #14** (line breaking), **UAX #9** (bidi)
- [ ] Hyphenation (sillabazione)
- [ ] Justification (allineamento giustificato)
- [ ] Line height avanzato da DOCX (`w:spacing/@lineRule`)

### ✅ Fase 4 - Styles & Theme mapping (Completata)
- [x] Parser `/word/styles.xml` con `w:basedOn`, `w:pStyle`, `w:rStyle`
- [x] Cascata completa: doc defaults → stile paragrafo → stile carattere → formattazione diretta
- [x] Theme font/color (`/word/theme/theme1.xml`) + `w:clrSchemeMapping` + tint/shade
- [x] Applicazione `ParagraphFormatting` (spacing, indentazioni, alignment) nel renderer
- [x] Nuovo comando `render` + campioni `styles-theme*.docx`
- [ ] Allineamento testo (left/center/right/justify)

### 🔮 Fase 4+ - Feature avanzate (Futuro)
- [ ] Colori testo e background (`w:color`, `w:shd`)
- [ ] Sottolineato, barrato (`w:u`, `w:strike`)
- [ ] Liste numerate e bullet (`w:numPr`)
- [ ] Tabelle (`w:tbl`)
- [ ] Header e footer (`w:hdr`, `w:ftr`)
- [ ] Immagini embedded (`w:drawing`, `w:pict`)
- [ ] Stili DOCX completi (`w:style`)
- [ ] Font embedding nel PDF
- [ ] Test di regressione visivi automatizzati

## Requisiti

- **.NET 9.0 SDK** o superiore
- **Supporto cross-platform**: macOS, Windows, Linux (grazie a SkiaSharp)

## Test

Il progetto attualmente include:
- **Smoke test manuale**: eseguire la demo e aprire `hello.pdf` in un viewer
- **Verifica testo selezionabile**: il testo nel PDF deve essere selezionabile (non immagine)
- **Verifica rendering**: caratteri speciali, legature e emoji devono renderizzare correttamente

Test automatizzati pianificati per le fasi successive.

## Licenza

TBD

## Riferimenti tecnici

- [SkiaSharp API Documentation](https://learn.microsoft.com/en-us/dotnet/api/skiasharp)
- [SkiaSharp.HarfBuzz](https://learn.microsoft.com/en-us/dotnet/api/skiasharp.harfbuzz)
- [Open XML SDK](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- [WordprocessingML Specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing)
- [Unicode UAX #14 (Line Breaking)](https://www.unicode.org/reports/tr14/)
- [Unicode UAX #9 (Bidirectional Algorithm)](https://www.unicode.org/reports/tr9/)

---

**DocxToPdf.Sdk** © 2025 - Motore di conversione DOCX → PDF in C# puro
